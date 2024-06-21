use comedy::com::ComRef;
use comedy::{com_call, HResult};
use winapi::shared::ntdef::SHORT;
use winapi::um::taskschd::{IDailyTrigger, ITrigger};
use crate::{bool_getter, string_getter, datetime_putter, get_repetition, short_getter};
use crate::ole_utils::BString;

pub struct DailyTrigger(pub ComRef<IDailyTrigger>);

impl DailyTrigger {
    bool_getter!(ITrigger::get_Enabled);
    get_repetition!(ITrigger::get_Repetition);
    string_getter!(ITrigger::get_StartBoundary);
    string_getter!(ITrigger::get_EndBoundary);

    short_getter!(IDailyTrigger::get_DaysInterval);
    string_getter!(IDailyTrigger::get_RandomDelay);

    datetime_putter!(IDailyTrigger::put_StartBoundary);


    // I'd like to have this only use the type-safe DateTime, but when copying it seems less
    // error-prone to use the string directly rather than try to parse it and then convert it back
    // to string.
    #[allow(non_snake_case)]
    pub fn put_StartBoundary_BString(&mut self, v: &BString) -> Result<(), HResult> {
        unsafe {
            com_call!(self.0, IDailyTrigger::put_StartBoundary(v.as_raw_ptr()))?;
        }
        Ok(())
    }


    #[allow(non_snake_case)]
    pub fn put_DaysInterval(&mut self, v: SHORT) -> Result<(), HResult> {
        unsafe {
            com_call!(self.0, IDailyTrigger::put_DaysInterval(v))?;
        }
        Ok(())
    }

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

        let interval = match self.get_DaysInterval() {
            Ok(val) => { format!(" [interval] {}", val)}
            Err(_) => {"".into()}
        };
        let random_delay = match self.get_RandomDelay() {
            Ok(val) => { format!(" [random_delay] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        format!("[DAILY]{}{}{}{}{}{}",enabled, start_boundary, end_boundary, interval, random_delay, repeat).trim().to_string()
    }
}
