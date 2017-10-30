#![recursion_limit = "1024"]
#[macro_use]
extern crate error_chain;
extern crate zip;
extern crate xml;

mod errors;

use std::fs;
use std::path::Path;
use std::io::{Read, Seek, SeekFrom, Cursor};
use std::collections::HashMap;
use std::str::FromStr;

use xml::reader::{EventReader, XmlEvent};

use errors::*;

// Skip BOM marker
fn skip_bom<R: Read + Seek>(reader: &mut R) -> Result<()> {
    let mut buffer = [0; 4];
    reader.seek(SeekFrom::Start(0))?;
    reader.read(&mut buffer[..])?;
    // UTF-8
    if buffer[0..3] == [0xef, 0xbb, 0xbf] {
        reader.seek(SeekFrom::Start(3))?;
    }
    // UTF-16BE
    else if buffer[0..2] == [0xfe, 0xff] {
        reader.seek(SeekFrom::Start(2))?;
    }
    // UTF-16LE
    else if buffer[0..2] == [0xff, 0xfe] {
        reader.seek(SeekFrom::Start(2))?;
    }
    // UTF-32BE
    else if buffer[0..4] == [0x00, 0x00, 0xfe, 0xff] {
        reader.seek(SeekFrom::Start(4))?;
    }
    // UTF-32LE
    else if buffer[0..4] == [0xff, 0xfe, 0x00, 0x00] {
        reader.seek(SeekFrom::Start(4))?;
    }
    Ok(())
}

#[derive(Debug, Clone)]
enum ValueType {
    String,
    Number,
}

#[derive(Debug, Clone)]
pub enum Value {
    String(String),
    Integer(i64),
    Float(f64),
    Empty,
}

struct Relation {
    target: String,
    kind: String,
}

struct Sheet {
    pub id: String,
    pub relation: String,
}

pub struct Cell {
    pub row: usize,
    pub column: usize,
    pub value: Value,
}

pub struct WorkSheet {
    pub cells: Vec<Cell>,
}

pub struct WorkBook {
    archive: zip::ZipArchive<fs::File>,
    strings: Vec<String>,
    relations: HashMap<String, Relation>,
    sheets: HashMap<String, Sheet>,
}

impl WorkBook {
    pub fn open<P: AsRef<Path>>(path: P) -> Result<WorkBook> {
        let file = fs::File::open(path)?;
        let archive = zip::ZipArchive::new(file)?;
        Ok(WorkBook {
            archive: archive,
            strings: vec![],
            sheets: HashMap::new(),
            relations: HashMap::new(),
            })
    }

    fn load_xml(&mut self, name: &str) -> Result<EventReader<Cursor<Vec<u8>>>> {
        let mut file = self.archive.by_name(name)?;
        // Unfortunaltely ZipFile does not support Seek and xml-rs does not support UTF BOM.
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer).unwrap();
        let mut file = Cursor::new(buffer);
        skip_bom(&mut file)?;
        Ok(EventReader::new(file))
    }

    fn load_relations<R: Read>(&mut self, reader: EventReader<R>) -> Result<()> {
        for ev in reader {
            match ev {
                Ok(XmlEvent::StartElement {name, attributes, ..}) => {
                    let mut rel_target = String::new();
                    let mut rel_id = String::new();
                    let mut rel_type = String::new();
                    if name.local_name == "Relationship" {
                        for attribute in attributes {
                            if attribute.name.local_name == "Target" {
                                rel_target = attribute.value;
                                if rel_target.starts_with("/") {
                                    rel_target.remove(0); 
                                }
                            }
                            else if attribute.name.local_name == "Id" {
                                rel_id = attribute.value;
                            }
                            else if attribute.name.local_name == "Type" {
                                rel_type = attribute.value;
                            }
                        }
                        self.relations.insert(rel_id, Relation {
                            target: rel_target,
                            kind: rel_type,
                            });
                    }
                }
                Err(error) => {
                    return Err(error.into());
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn load_relationships(&mut self) -> Result<()> {
        for i in 0..self.archive.len() {
            let mut file_name = String::new();
            {
                let file = self.archive.by_index(i)?;
                file_name.push_str(file.name());
            }
            if file_name.ends_with(".rels") {
                let reader = self.load_xml(&file_name)?;
                self.load_relations(reader)?;
            }
        }
        Ok(())
    }

    /// Read shared string list
    fn load_shared_strings(&mut self) -> Result<()> {
        let mut path = String::new();
        for (_, relation) in self.relations.iter() {
            if relation.kind.ends_with("/sharedStrings") {
                path = relation.target.clone();
                break;
            }
        }
        let reader = self.load_xml(&path)?;
        let mut store = false;
        for ev in reader {
            match ev {
                Ok(XmlEvent::StartElement {name, ..}) => {
                    if name.local_name == "t" {
                        store = true;
                    }
                    else {
                        store = false;
                    } 
                }
                Ok(XmlEvent::Characters(text)) => {
                    if store {
                        self.strings.push(text);
                    }
                }
                Err(error) => {
                    return Err(error.into());
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn load_workbook(&mut self) -> Result<()> {
        let mut path = String::new();
        for (_, relation) in self.relations.iter() {
            if relation.kind.ends_with("/officeDocument") {
                path = relation.target.clone();
                break;
            }
        }
        let reader = self.load_xml(&path)?;
        for ev in reader {
            match ev {
                Ok(XmlEvent::StartElement {name, attributes, ..}) => {
                    if name.local_name == "sheet" {
                        let mut sheet_name = String::new();
                        let mut sheed_id = String::new();
                        let mut sheet_relation_id = String::new();
                        for attribute in attributes {
                            if attribute.name.local_name == "name" {
                                sheet_name = attribute.value;
                            }
                            else if attribute.name.local_name == "sheetId" {
                                sheed_id = attribute.value;
                            }
                            else if attribute.name.local_name == "id" {
                                sheet_relation_id = attribute.value;
                            }
                        }
                        self.sheets.insert(sheet_name,
                            Sheet {
                                relation: sheet_relation_id,
                                id: sheed_id,
                                });
                    }
                }
                Err(error) => {
                    return Err(error.into());
                }
                _ => {}
            }
        }
        Ok(())
    }
    
    pub fn list_worksheet(&self) -> Result<Vec<String>> {
        let mut worksheets = Vec::new();
        for name in self.sheets.keys() {
            worksheets.push(name.clone());
        }
        Ok(worksheets)
    }
    
    pub fn load_worksheet(&mut self, name: &str) -> Result<WorkSheet> {
        let ref sheet_relation = self.sheets[name].relation.clone();
        let ref target = self.relations[sheet_relation].target.clone();
        let reader = self.load_xml(target)?;
        let mut row = 0usize;
        let mut column = 0usize;
        let mut kind = ValueType::String;
        let mut capture_value = false;
        let mut sheet = WorkSheet { cells: Vec::new() }; 
        for ev in reader {
            match ev {
                Ok(XmlEvent::StartElement {name, attributes, ..}) => {
                    capture_value = false;
                    if name.local_name == "row" {
                        for attribute in attributes.iter() {
                            if attribute.name.local_name == "r" {
                                row = usize::from_str(&attribute.value)?;
                            }
                        }
                        column = 0;
                    }
                    else if name.local_name == "c" {
                        for attribute in attributes.iter() {
                            if attribute.name.local_name == "t" {
                                if attribute.value == "s" {
                                    kind = ValueType::String;
                                }
                                else if attribute.value == "n" {
                                    kind = ValueType::Number;
                                }
                            }
                        }
                        column += 1;
                    }
                    else if name.local_name == "v" {
                        capture_value = true;
                    }
                }
                Ok(XmlEvent::Characters(text)) => {
                    if capture_value {
                        let value = match kind {
                            ValueType::String => {
                                let index = usize::from_str(&text)?;
                                Value::String(self.strings[index].clone())
                            },
                            ValueType::Number => {
                                match i64::from_str(&text) {
                                    Ok(value) => { Value::Integer(value) }
                                    Err(_) => {
                                        let value = f64::from_str(&text)?;
                                        Value::Float(value)
                                    }
                                }
                            }
                        };
                        sheet.cells.push(Cell { row: row, column: column, value: value });
                    }
                }
                Err(error) => {
                    return Err(error.into());
                }
                _ => {}
            }
        }
        Ok(sheet)
    }

    pub fn load(&mut self) -> Result<()> {
        self.load_relationships()?;
        self.load_workbook()?;
        self.load_shared_strings()?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
    }
}
