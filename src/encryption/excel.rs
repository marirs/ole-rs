use crate::{
    encryption::{DocumentType, EncryptionHandler},
    OleFile,
};
use std::collections::HashMap;

lazy_static! {
    pub static ref NAME_TO_RECORD_NUM_MAP: HashMap<&'static str, u16> = {
        HashMap::from([
            ("AlRuns", 4176),
            ("Area", 4122),
            ("AreaFormat", 4106),
            ("Array", 545),
            ("AttachedLabel", 4108),
            ("AutoFilter", 158),
            ("AutoFilter12", 2174),
            ("AutoFilterInfo", 157),
            ("AxcExt", 4194),
            ("AxesUsed", 4166),
            ("Axis", 4125),
            ("AxisLine", 4129),
            ("AxisParent", 4161),
            ("BCUsrs", 407),
            ("BOF", 2057),
            ("BRAI", 4177),
            ("Backup", 64),
            ("Bar", 4119),
            ("Begin", 4147),
            ("BigName", 1048),
            ("BkHim", 233),
            ("Blank", 513),
            ("BookBool", 218),
            ("BookExt", 2147),
            ("BoolErr", 517),
            ("BopPop", 4193),
            ("BopPopCustom", 4199),
            ("BottomMargin", 41),
            ("BoundSheet8", 133),
            ("BuiltInFnGroupCount", 156),
            ("CF", 433),
            ("CF12", 2170),
            ("CFEx", 2171),
            ("CRN", 90),
            ("CUsr", 401),
            ("CalcCount", 12),
            ("CalcDelta", 16),
            ("CalcIter", 17),
            ("CalcMode", 13),
            ("CalcPrecision", 14),
            ("CalcRefMode", 15),
            ("CalcSaveRecalc", 95),
            ("CatLab", 2134),
            ("CatSerRange", 4128),
            ("CbUsr", 402),
            ("CellWatch", 2156),
            ("Chart", 4098),
            ("Chart3DBarShape", 4191),
            ("Chart3d", 4154),
            ("ChartFormat", 4116),
            ("ChartFrtInfo", 2128),
            ("ClrtClient", 4188),
            ("CodeName", 442),
            ("CodePage", 66),
            ("ColInfo", 125),
            ("Compat12", 2188),
            ("CompressPictures", 2203),
            ("CondFmt", 432),
            ("CondFmt12", 2169),
            ("Continue", 60),
            ("ContinueBigName", 1084),
            ("ContinueFrt", 2066),
            ("ContinueFrt11", 2165),
            ("ContinueFrt12", 2175),
            ("Country", 140),
            ("CrErr", 2149),
            ("CrtLayout12", 2205),
            ("CrtLayout12A", 2215),
            ("CrtLine", 4124),
            ("CrtLink", 4130),
            ("CrtMlFrt", 2206),
            ("CrtMlFrtContinue", 2207),
            ("DBCell", 215),
            ("DBQueryExt", 2051),
            ("DCon", 80),
            ("DConBin", 437),
            ("DConName", 82),
            ("DConRef", 81),
            ("DConn", 2166),
            ("DSF", 353),
            ("DVal", 434),
            ("DXF", 2189),
            ("Dat", 4195),
            ("DataFormat", 4102),
            ("DataLabExt", 2154),
            ("DataLabExtContents", 2155),
            ("Date1904", 34),
            ("DbOrParamQry", 220),
            ("DefColWidth", 85),
            ("DefaultRowHeight", 549),
            ("DefaultText", 4132),
            ("Dimensions", 512),
            ("DocRoute", 184),
            ("DropBar", 4157),
            ("DropDownObjIds", 2164),
            ("Dv", 446),
            ("DxGCol", 153),
            ("EOF", 10),
            ("End", 4148),
            ("EndBlock", 2131),
            ("EndObject", 2133),
            ("EntExU2", 450),
            ("Excel9File", 448),
            ("ExtSST", 255),
            ("ExtString", 2052),
            ("ExternName", 35),
            ("ExternSheet", 23),
            ("Fbi", 4192),
            ("Fbi2", 4200),
            ("Feat", 2152),
            ("FeatHdr", 2151),
            ("FeatHdr11", 2161),
            ("Feature11", 2162),
            ("Feature12", 2168),
            ("FileLock", 405),
            ("FilePass", 47),
            ("FileSharing", 91),
            ("FilterMode", 155),
            ("FnGroupName", 154),
            ("FnGrp12", 2200),
            ("Font", 49),
            ("FontX", 4134),
            ("Footer", 21),
            ("ForceFullCalculation", 2211),
            ("Format", 1054),
            ("Formula", 6),
            ("Frame", 4146),
            ("FrtFontList", 2138),
            ("FrtWrapper", 2129),
            ("GUIDTypeLib", 2199),
            ("GelFrame", 4198),
            ("GridSet", 130),
            ("Guts", 128),
            ("HCenter", 131),
            ("HFPicture", 2150),
            ("HLink", 440),
            ("HLinkTooltip", 2048),
            ("Header", 20),
            ("HeaderFooter", 2204),
            ("HideObj", 141),
            ("HorizontalPageBreaks", 27),
            ("IFmtRecord", 4174),
            ("Index", 523),
            ("InterfaceEnd", 226),
            ("InterfaceHdr", 225),
            ("Intl", 97),
            ("LPr", 152),
            ("LRng", 351),
            ("Label", 516),
            ("LabelSst", 253),
            ("Lbl", 24),
            ("LeftMargin", 38),
            ("Legend", 4117),
            ("LegendException", 4163),
            ("Lel", 441),
            ("Line", 4120),
            ("LineFormat", 4103),
            ("List12", 2167),
            ("MDB", 2186),
            ("MDTInfo", 2180),
            ("MDXKPI", 2185),
            ("MDXProp", 2184),
            ("MDXSet", 2183),
            ("MDXStr", 2181),
            ("MDXTuple", 2182),
            ("MTRSettings", 2202),
            ("MarkerFormat", 4105),
            ("MergeCells", 229),
            ("Mms", 193),
            ("MsoDrawing", 236),
            ("MsoDrawingGroup", 235),
            ("MsoDrawingSelection", 237),
            ("MulBlank", 190),
            ("MulRk", 189),
            ("NameCmt", 2196),
            ("NameFnGrp12", 2201),
            ("NamePublish", 2195),
            ("Note", 28),
            ("Number", 515),
            ("ObNoMacros", 445),
            ("ObProj", 211),
            ("Obj", 93),
            ("ObjProtect", 99),
            ("ObjectLink", 4135),
            ("OleDbConn", 2058),
            ("OleObjectSize", 222),
            ("PLV", 2187),
            ("Palette", 146),
            ("Pane", 65),
            ("Password", 19),
            ("PhoneticInfo", 239),
            ("PicF", 4156),
            ("Pie", 4121),
            ("PieFormat", 4107),
            ("PivotChartBits", 2137),
            ("PlotArea", 4149),
            ("PlotGrowth", 4196),
            ("Pls", 77),
            ("Pos", 4175),
            ("PrintGrid", 43),
            ("PrintRowCol", 42),
            ("PrintSize", 51),
            ("Prot4Rev", 431),
            ("Prot4RevPass", 444),
            ("Protect", 18),
            ("Qsi", 429),
            ("QsiSXTag", 2050),
            ("Qsif", 2055),
            ("Qsir", 2054),
            ("RK", 638),
            ("RRAutoFmt", 331),
            ("RRDChgCell", 315),
            ("RRDConflict", 338),
            ("RRDDefName", 339),
            ("RRDHead", 312),
            ("RRDInfo", 406),
            ("RRDInsDel", 311),
            ("RRDInsDelBegin", 336),
            ("RRDInsDelEnd", 337),
            ("RRDMove", 320),
            ("RRDMoveBegin", 334),
            ("RRDMoveEnd", 335),
            ("RRDRenSheet", 318),
            ("RRDRstEtxp", 340),
            ("RRDTQSIF", 2056),
            ("RRDUserView", 428),
            ("RRFormat", 330),
            ("RRInsertSh", 333),
            ("RRSort", 319),
            ("RRTabId", 317),
            ("Radar", 4158),
            ("RadarArea", 4160),
            ("RealTimeData", 2067),
            ("RecalcId", 449),
            ("RecipName", 185),
            ("RefreshAll", 439),
            ("RichTextStream", 2214),
            ("RightMargin", 39),
            ("Row", 520),
            ("SBaseRef", 4168),
            ("SCENARIO", 175),
            ("SIIndex", 4197),
            ("SST", 252),
            ("SXAddl", 2148),
            ("SXDB", 198),
            ("SXDBB", 200),
            ("SXDBEx", 290),
            ("SXDI", 197),
            ("SXDtr", 206),
            ("SXEx", 241),
            ("SXFDB", 199),
            ("SXFDBType", 443),
            ("SXFormula", 259),
            ("SXInt", 204),
            ("SXLI", 181),
            ("SXNum", 201),
            ("SXPI", 182),
            ("SXPIEx", 2062),
            ("SXPair", 248),
            ("SXRng", 216),
            ("SXStreamID", 213),
            ("SXString", 205),
            ("SXTBRGIITM", 209),
            ("SXTH", 2061),
            ("SXTbl", 208),
            ("SXVDEx", 256),
            ("SXVDTEx", 2063),
            ("SXVI", 178),
            ("SXVS", 227),
            ("SXViewEx", 2060),
            ("SXViewEx9", 2064),
            ("SXViewLink", 2136),
            ("Scatter", 4123),
            ("ScenMan", 174),
            ("ScenarioProtect", 221),
            ("Scl", 160),
            ("Selection", 29),
            ("SerAuxErrBar", 4187),
            ("SerAuxTrend", 4171),
            ("SerFmt", 4189),
            ("SerParent", 4170),
            ("SerToCrt", 4165),
            ("Series", 4099),
            ("SeriesList", 4118),
            ("SeriesText", 4109),
            ("Setup", 161),
            ("ShapePropsStream", 2212),
            ("SheetExt", 2146),
            ("ShrFmla", 1212),
            ("ShtProps", 4164),
            ("Sort", 144),
            ("SortData", 2197),
            ("StartBlock", 2130),
            ("StartObject", 2132),
            ("String", 519),
            ("Style", 659),
            ("StyleExt", 2194),
            ("SupBook", 430),
            ("Surf", 4159),
            ("SxBool", 202),
            ("SxDXF", 244),
            ("SxErr", 203),
            ("SxFilt", 242),
            ("SxFmla", 249),
            ("SxFormat", 251),
            ("SxIsxoper", 217),
            ("SxItm", 245),
            ("SxIvd", 180),
            ("SxName", 246),
            ("SxNil", 207),
            ("SxRule", 240),
            ("SxSelect", 247),
            ("SxTbpg", 210),
            ("SxView", 176),
            ("Sxvd", 177),
            ("Sync", 151),
            ("Table", 566),
            ("TableStyle", 2191),
            ("TableStyleElement", 2192),
            ("TableStyles", 2190),
            ("Template", 96),
            ("Text", 4133),
            ("TextPropsStream", 2213),
            ("Theme", 2198),
            ("Tick", 4126),
            ("TopMargin", 40),
            ("TxO", 438),
            ("TxtQry", 2053),
            ("Uncalced", 94),
            ("Units", 4097),
            ("UserBView", 425),
            ("UserSViewBegin", 426),
            ("UserSViewBegin_Chart", 426),
            ("UserSViewEnd", 427),
            ("UsesELFs", 352),
            ("UsrChk", 408),
            ("UsrExcl", 404),
            ("UsrInfo", 403),
            ("VCenter", 132),
            ("ValueRange", 4127),
            ("VerticalPageBreaks", 26),
            ("WOpt", 2059),
            ("WebPub", 2049),
            ("WinProtect", 25),
            ("Window1", 61),
            ("Window2", 574),
            ("WriteAccess", 92),
            ("WriteProtect", 134),
            ("WsBool", 129),
            ("XCT", 89),
            ("XF", 224),
            ("XFCRC", 2172),
            ("XFExt", 2173),
            ("YMult", 2135),
        ])
    };
}

struct BIFFSTream<'a> {
    data: &'a [u8],
    iterator_position: Option<usize>,
}

impl<'a> BIFFSTream<'a> {
    pub fn new(stream: &'a [u8]) -> Self {
        Self {
            data: stream,
            iterator_position: None,
        }
    }

    pub fn has_record(&mut self, target: u16) -> bool {
        self.reset();
        self.into_iter().any(|item| item.num == target)
    }

    pub fn skip_to(&mut self, target: u16) -> Option<BiffItem> {
        self.reset();
        self.into_iter().find(|item| item.num == target)
    }

    pub fn reset(&mut self) {
        self.iterator_position = None;
    }
}

struct BiffItem<'a> {
    pub num: u16,
    pub size: u16,
    pub data: &'a [u8],
}

impl<'a> Iterator for BIFFSTream<'a> {
    type Item = BiffItem<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let position = self.iterator_position.unwrap_or(0);

        let len = self.data.len();
        let end_of_position_slice = position + 4;
        if end_of_position_slice >= len {
            return None;
        }
        let h = &self.data[position..end_of_position_slice];
        let num = u16::from_le_bytes([h[0], h[1]]);
        let size = u16::from_le_bytes([h[2], h[3]]);
        let end = end_of_position_slice + size as usize;
        self.iterator_position = Some(end);
        Some(BiffItem {
            num,
            size,
            data: &self.data[end_of_position_slice..end],
        })
    }
}

pub(crate) struct ExcelEncryptionHandler<'a> {
    ole_file: &'a OleFile,
    stream_name: String,
}

impl<'a> EncryptionHandler<'a> for ExcelEncryptionHandler<'a> {
    fn doc_type(&self) -> DocumentType {
        DocumentType::Excel
    }

    fn is_encrypted(&self) -> bool {
        let workbook_stream = self
            .ole_file
            .open_stream(&[self.stream_name.as_str()])
            .expect("unable to open workbook?");
        let workbook = BIFFSTream::new(&workbook_stream);
        let first = workbook.into_iter().next().expect("must have first item");
        assert_eq!(&first.num, NAME_TO_RECORD_NUM_MAP.get("BOF").unwrap());
        let mut workbook = BIFFSTream::new(&workbook_stream);
        match workbook.skip_to(*NAME_TO_RECORD_NUM_MAP.get("FilePass").unwrap()) {
            Some(item) => {
                match &item.data[0..2] {
                    [0x01, 0x00] => {
                        //RC4
                        true
                    },
                    [0x00, 0x00] => {
                        // XOR Obfuscation unsupported
                        false
                    },
                    _ => {
                        //anything else is not encrypted
                        false
                    }
                }
            },
            None => {
                false
            },
        }
    }

    fn new(ole_file: &'a OleFile, stream_name: String) -> Self {
        Self {
            ole_file,
            stream_name,
        }
    }
}
