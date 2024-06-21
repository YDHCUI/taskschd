use comedy::com::ComRef;
use winapi::um::taskschd;
use winapi::um::taskschd::{IPrincipal, TASK_LOGON_TYPE, TASK_RUNLEVEL};
use crate::{string_getter, short_getter};

pub struct Principal(pub ComRef<IPrincipal>);

impl Principal {
    string_getter!(IPrincipal::get_DisplayName);
    string_getter!(IPrincipal::get_UserId);


    #[allow(non_snake_case)]
    pub fn get_logon_type(&mut self) -> Result<TASK_LOGON_TYPE, comedy::HResult> {
        unsafe {
            let mut i = 0;
            comedy::com_call!(self.0, IPrincipal::get_LogonType(&mut i))?;
            Ok(i)
        }
    }

    #[allow(non_snake_case)]
    pub fn get_run_level(&mut self) -> Result<TASK_RUNLEVEL, comedy::HResult> {
        unsafe {
            let mut i = 0;
            comedy::com_call!(self.0, IPrincipal::get_RunLevel(&mut i))?;
            Ok(i)
        }
    }

    pub fn to_string(&mut self) -> String {
        let name = match self.get_DisplayName() {
            Ok(val) => { format!(" [name] {}", val)}
            Err(_) => {"".into()}
        };
        let logon_type = match self.get_logon_type() {
            Ok(val) => {
                let logon_type = match val {
                    taskschd::TASK_LOGON_NONE => "none",
                    taskschd::TASK_LOGON_PASSWORD => "password",
                    taskschd::TASK_LOGON_S4U => "s4u",
                    taskschd::TASK_LOGON_INTERACTIVE_TOKEN => "interactive_token",
                    taskschd::TASK_LOGON_GROUP => "group",
                    taskschd::TASK_LOGON_SERVICE_ACCOUNT => "service_account",
                    taskschd::TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD => "interactive_token_or_password",
                    _ => "",
                };
                format!(" [logon_type] {}", logon_type)
            }
            Err(_) => {"".into()}
        };
        let run_level = match self.get_run_level() {
            Ok(val) => {
                let logon_type = match val {
                    taskschd::TASK_RUNLEVEL_LUA => "least",
                    taskschd::TASK_RUNLEVEL_HIGHEST => "highest",
                    _ => "",
                };
                format!(" [run_level] {}", logon_type)
            }
            Err(_) => {"".into()}
        };
        format!("[PRINCIPAL]{}{}{}",name,logon_type,run_level)

    }


}