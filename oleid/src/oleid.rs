use std::fmt::{Debug, Formatter};
use ole::{OleFile};
use ole::ftype::OleFileType;

/// Constants for risk values.
#[derive(Debug, Clone)]
pub enum Risk {
    HIGH,
    MEDIUM,
    LOW,
    NONE,
    INFO,
    UNKNOWN,
    /// If a check triggered an unexpected error.
    ERROR
}

/// Piece of information of an `OleID` object.
/// Contains an ID, value, type, name and description. No other functionality.
#[derive(Clone)]
pub struct Indicator {
    id: String,
    value: Option<String>,
    // Not sure we need this
    _type: String,
    name: Option<String>,
    description: Option<String>,
    risk: Risk,
    hide_if_false: bool
}

impl Indicator {
    pub fn new(id: &str, 
               value: Option<&str>, 
               _type: &str, 
               name: Option<&str>, 
               description: Option<&str>, 
               risk: Risk, 
               hide_if_false: bool) -> Self {
        Indicator {
            id: id.to_string(),
            value: value.map(|x| x.to_string()),
            _type: _type.to_string(),
            name: name.map(|x| x.to_string()),
            description: description.map(|x| x.to_string()),
            risk,
            hide_if_false
        }
    }
}

impl Debug for Indicator {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Indicator")
            .field("name", &self.name)
            .field("value", &self.value)
            .field("risk", &self.risk)
            .field("description", &self.description)
            .finish()
    }
}

/// Summary of information about an OLE file (and a few other MS Office formats)
/// Call `OleID::check()` to gather all info on a given file or run one
/// of the `check_` functions to just get a specific piece of info.
pub struct OleId {
    indicators: Vec<Indicator>,
    suminfo_data: Option<String>,
    ole: Option<OleFile>
}

impl OleId {
    /// Create an OleID object.  
    ///         This does not run any checks yet nor open the file.
    ///         Can either give just a filename (as str), so OleID will check whether
    ///         that is a valid OLE file and create a `OleFile`
    ///         object for it.
    ///         If filename is given, only `OleID::check` opens the file. Other
    ///         functions will return None
    pub fn new(filename: &str) -> Self {
        OleId {
            indicators: Vec::new(),
            suminfo_data: None,
            ole: match OleFile::from_file_blocking(filename){
                Ok(t) => Some(t),
                _=> {
                    panic!("Could not parse the provided file.");
                }
            }
        }
    }
    
    /// Open file and run all checks on it.
    /// returns: list of all `Indicator`s created
    pub fn check(&mut self) -> Vec<Indicator> {
        // We have a value so far but check the value of the ole file object available just to be sure.
        let file_type = match self.ole.as_ref().cloned(){
            Some(t) => t.file_type,
            _=> {
                panic!("The ole file is invalid.");
            }
        };
        let description = match file_type {
            OleFileType::Generic => Some("Unrecognized OLE file."),
            _=> None
        };
        let filetype_indicator = Indicator::new("FType", Some(format!("{:?}", file_type).as_str()), "String", Some("File format"), description, Risk::INFO, true);
        self.indicators.push(filetype_indicator);
        
        self.check_encrypted();
        self.check_macros();
        self.check_external_relationships();
        self.check_object_pool();
        self.check_flash();
        self.indicators.clone()
    }

    /// Check whether this file is encrypted.
    pub fn check_encrypted(&mut self) -> Indicator {
        let mut encrypted_indicator = Indicator::new("Encrypted", None, "Bool", Some("Encrypted"), Some("The file is not encrypted"), Risk::NONE, false);
        if self.ole.as_ref().cloned().unwrap().encrypted {
            encrypted_indicator.value = Some("True".to_string());
            encrypted_indicator.risk = Risk::LOW;
            encrypted_indicator.description = Some("The file is encrypted. It may be decrypted with msoffcrypto-tool".to_string());
        }
        self.indicators.push(encrypted_indicator.clone());
        encrypted_indicator
    }

    /// Check whether this file contains macros (VBA and XLM/Excel 4).
    pub fn check_macros(&mut self) {
        let macros_indicator = Indicator::new("vba", Some("No"), "String", Some("VBA Macros"), Some("This file does not contain VBA macros."), Risk::NONE, false);
        self.indicators.push(macros_indicator.clone());
        let xlm_indicator = Indicator::new("xlm", Some("No"), "String", Some("XLM Macros"), Some("This file does not contain Excel 4/XLM macros."), Risk::NONE, false);
        self.indicators.push(xlm_indicator.clone());
        // Check XLM Macros only in excel files
        if self.ole.as_ref().cloned().unwrap().is_excel() {
            // TODO: Hook up with the VBA Parser of the VBA module
        }
    }

    ///  Check whether this file has external relationships (remote template, OLE object, etc).
    pub fn check_external_relationships(&mut self) -> Indicator{
        let external_relations_indicator = Indicator::new("ExternalRelations", None, "Int", Some("External Relationships"), Some("External relationships such as remote templates, remote OLE objects, etc"), Risk::NONE, false);
        self.indicators.push(external_relations_indicator.clone());
        // TODO: This check is only for openxml files.
        // match self.ole.as_ref().cloned().unwrap().file_type {
        //     OleFileType::
        // }
        external_relations_indicator
    }

    /// Check whether this file contains an ObjectPool stream.
    /// Such a stream would be a strong indicator for embedded objects or files.
    pub fn check_object_pool(&mut self) -> Indicator {
        let mut object_pool_indicator = Indicator::new("ObjectPool", None, "Int", Some("Object Pool"), Some("Contains an ObjectPool stream, very likely to contain embedded OLE objects or files. Use oleobj to check it."), Risk::NONE, false);
        if self.ole.as_ref().cloned().unwrap().list_streams().contains(&"ObjectPool".to_string()) {
            object_pool_indicator.value = Some("True".to_string());
            object_pool_indicator.risk = Risk::LOW;
        }
        self.indicators.push(object_pool_indicator.clone());
        object_pool_indicator
    }

    /// Check whether this file contains flash objects
    pub fn check_flash(&mut self) -> Indicator {
        let mut flash_indicator = Indicator::new("Flash", Some("0"), "Int", Some("Flash Objects"), Some("Number of embedded Flash objects (SWF files) detected in OLE streams. Not 100% accurate, there may be false positives."), Risk::NONE, false);
        self.indicators.push(flash_indicator.clone());
        let found = detect_flash(self.ole.as_ref().cloned().unwrap().directory_stream_data);
        let val = flash_indicator.value.as_ref().cloned().unwrap().parse::<i32>().unwrap();
        let new_val = val + found.len() as i32;
        flash_indicator.value = Some(new_val.to_string());
        flash_indicator
    }

    /// Helper function: returns an indicator if present (or None)
    pub fn get_indicator(&self, indicator_id: &str) -> Option<Indicator> {
        return match self.indicators.iter().find(|indicator| indicator.id == indicator_id) {
            Some(t) => Some(t.clone()),
            _=> None
        }
    }
}

pub fn detect_flash(stream_data: Vec<u8>) -> Vec<String> {
    vec![]
}