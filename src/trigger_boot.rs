use comedy::com::ComRef;
use winapi::um::taskschd::{IBootTrigger, ITrigger};
use crate::{bool_getter, string_getter, get_repetition};

pub struct BootTrigger(pub ComRef<IBootTrigger>);

impl BootTrigger {
    bool_getter!(ITrigger::get_Enabled);
    get_repetition!(ITrigger::get_Repetition);
    string_getter!(ITrigger::get_StartBoundary);
    string_getter!(ITrigger::get_EndBoundary);


    string_getter!(IBootTrigger::get_Delay);


    pub fn to_string(&mut self) -> String {
        let repeat = match self.get_Repetition() {
            Ok(val) => { format!(" {}", val.to_string().trim())}
            Err(_) => {"".into()}
        };
        let enabled = match self.get_Enabled() {
            Ok(val) => { if val {" [ENABLED]".to_string()} else {" DISABLED".to_string()} },
            Err(_) => {"".to_string()}
        };

        let start = match self.get_StartBoundary() {
            Ok(val) => { format!(" [start_boundary] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let end = match self.get_EndBoundary() {
            Ok(val) => { format!(" [end_boundary] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        let delay = match self.get_Delay() {
            Ok(val) => { format!(" [delay] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        format!("[BOOT]{}{}{}{}{}",enabled, start, end, delay,repeat).trim().to_string()
    }
}