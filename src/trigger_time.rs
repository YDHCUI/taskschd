use comedy::com::ComRef;
use winapi::um::taskschd::{ ITimeTrigger, ITrigger};
use crate::{bool_getter, bstring_getter, get_repetition};
pub struct TimeTrigger(pub ComRef<ITimeTrigger>);

impl TimeTrigger {
    bool_getter!(ITrigger::get_Enabled);
    get_repetition!(ITrigger::get_Repetition);
    bstring_getter!(ITrigger::get_StartBoundary);
    bstring_getter!(ITrigger::get_EndBoundary);


    bstring_getter!(ITimeTrigger::get_RandomDelay);

    pub fn to_string(&mut self) -> String {
        let enabled = match self.get_Enabled() {
            Ok(val) => { if val {" [ENABLED]".to_string()} else {" DISABLED".to_string()} },
            Err(_) => {"".to_string()}
        };
        let repeat = match self.get_Repetition() {
            Ok(val) => { format!(" {}", val.to_string().trim())}
            Err(_) => {"".into()}
        };
        let start_boundary = match self.get_StartBoundary() {
            Ok(val) => { format!(" [start_boundary] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let end_boundary = match self.get_EndBoundary() {
            Ok(val) => { format!(" [end_boundary] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let delay = match self.get_RandomDelay() {
            Ok(val) => { format!(" [random_delay] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        format!("[TIME]{}{}{}{}{}",enabled, start_boundary, end_boundary, delay,repeat).trim().to_string()
    }
}