use std::ops::RangeFrom;

pub trait StringUtils {
    fn substring(&self, range: RangeFrom<usize>) -> String;
}

impl StringUtils for String {
    fn substring(&self, range: RangeFrom<usize>) -> String {
        let chars = self.chars().map(|x| x.to_string()).collect::<Vec<_>>();
        let mut substring = Vec::new();
        for char_index in range.start..chars.len() {
            let cur_char = chars[char_index - 1].clone();
            substring.push(cur_char);
        }
        substring.join("")
    }
}
