
use std::num;
use std::str;
use std::string;
use std::io;

use xml;
use zip;

error_chain! {
    types {
        Error, ErrorKind, ChainErr, Result;
    }

    foreign_links {
        num::ParseIntError, ParseInt;
        num::ParseFloatError, ParseFloat;
        str::Utf8Error, Utf8;
        string::FromUtf8Error, FromUtf8;
        io::Error, Io;
        zip::result::ZipError, Zip;
        xml::reader::Error, XmlRead;
    }
}
