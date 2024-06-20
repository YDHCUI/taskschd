use comedy::com::ComRef;
use winapi::um::taskschd::{ ITrigger, IMonthlyTrigger};
use crate::{bool_getter, bstring_getter, get_repetition, long_getter, short_getter};

pub struct MonthlyTrigger(pub ComRef<IMonthlyTrigger>);

impl MonthlyTrigger {
    bool_getter!(ITrigger::get_Enabled);
    get_repetition!(ITrigger::get_Repetition);
    bstring_getter!(ITrigger::get_StartBoundary);
    bstring_getter!(ITrigger::get_EndBoundary);

    bstring_getter!(IMonthlyTrigger::get_RandomDelay);

    long_getter!(IMonthlyTrigger::get_DaysOfMonth);
    short_getter!(IMonthlyTrigger::get_MonthsOfYear);
    // bool_getter!(IMonthlyTrigger::get_RunOnLastDayOfMonth);

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

        let days_month = match self.get_DaysOfMonth() {
            Ok(val) => { format!(" [days_month] {}", val.to_string())}
            Err(_) => {"".into()}
        };

        let months_of_year = match self.get_MonthsOfYear() {
            Ok(val) => { format!(" [months_of_year] {}", val.to_string())}
            Err(_) => {"".into()}
        };


        let random_delay = match self.get_RandomDelay() {
            Ok(val) => { format!(" [random_delay] {}", val.to_string())}
            Err(_) => {"".into()}
        };
        format!("[MONTHLY]{}{}{}{}{}{}{}",enabled, start_boundary, end_boundary, days_month,months_of_year, random_delay, repeat).trim().to_string()
    }

}