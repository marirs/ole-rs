use std::fmt::{Debug, Formatter};
use std::fs;

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
    r#type: bool,
    name: Option<String>,
    description: Option<String>,
    risk: Risk,
    hide_if_false: bool
}

impl Indicator {
    pub fn new(id: String, 
               value: Option<&str>, 
               r#type: bool, 
               name: Option<&str>, 
               description: Option<&str>, 
               risk: Risk, 
               hide_if_false: bool) -> Self {
        Indicator {
            id,
            value: value.map(|x| x.to_string()),
            r#type,
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
    suminfo_data: Option<String>
}

impl OleId {
    /// Create an OleID object.  
    ///         This does not run any checks yet nor open the file.
    ///         Can either give just a filename (as str), so OleID will check whether
    ///         that is a valid OLE file and create a :py:class:`olefile.OleFileIO`
    ///         object for it. Or you can give an already opened
    ///         :py:class:`OleFileIO` as argument to avoid re-opening (e.g. if
    ///         called from other oletools).
    ///         If filename is given, only `OleID::check` opens the file. Other
    ///         functions will return None
    pub fn new(filename: &str) -> Self {
        // Read the file.
        let contents = fs::read(filename).expect("Could not read the file.");
        OleId {
            indicators: Vec::new(),
            suminfo_data: None
        }
    }
    
    /// Open file and run all checks on it.
    /// returns: list of all `Indicator`s created
    pub fn check(&mut self) -> Vec<Indicator> {
        // self.check_properties();
        // self.check_encrypted();
        // self.check_macros();
        // self.check_external_relationships();
        // self.check_object_pool();
        // self.check_flash();
        self.indicators.clone()
    }

    /// Helper function: returns an indicator if present (or None)
    pub fn get_indicator(&self, indicator_id: &str) -> Option<Indicator> {
        return match self.indicators.iter().find(|indicator| indicator.id == indicator_id) {
            Some(t) => Some(t.clone()),
            _=> None
        }
    }
}