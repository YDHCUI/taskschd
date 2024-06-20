use comedy::com::ComRef;
use winapi::um::taskschd::IRepetitionPattern;
use crate::{bool_getter, bstring_getter};

pub struct RepetitionPattern(pub ComRef<IRepetitionPattern>);

impl RepetitionPattern {
    bstring_getter!(IRepetitionPattern::get_Interval);
    bstring_getter!(IRepetitionPattern::get_Duration);
    bool_getter!(IRepetitionPattern::get_StopAtDurationEnd);

    pub fn to_string(&mut self) -> String {
        let interval = match self.get_Interval() {
            Ok(val) => { format!(" [interval] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let duration = match self.get_Duration() {
            Ok(val) => { format!(" [duration] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let enabled = match self.get_StopAtDurationEnd() {
            Ok(val) => { if val {" [STOP AT END]".to_string()} else {"".to_string()} },
            Err(_) => {"".to_string()}
        };
        if interval.is_empty() && duration.is_empty() && enabled.is_empty() {
            return "".into()
        }
        format!("<[REPEAT]{}{}{}>",interval,duration,enabled)
    }
}
