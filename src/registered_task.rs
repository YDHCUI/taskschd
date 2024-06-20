use std::ffi::{OsStr};
use comedy::com::ComRef;
use comedy::{com_call, com_call_getter, HResult};
use winapi::ctypes::c_double;
use winapi::shared::wtypes::DATE;
use winapi::um::taskschd;
use winapi::um::taskschd::{IRegisteredTask, IRunningTask, TASK_STATE};
use crate::ole_utils::{BString, date_to_datetime, OptionBstringExt};
use crate::task_definition::TaskDefinition;
use crate::{bstring_getter, long_getter, try_to_bstring};
pub struct RegisteredTask(pub ComRef<IRegisteredTask>);

impl RegisteredTask {
    bstring_getter!(IRegisteredTask::get_Name);
    bstring_getter!(IRegisteredTask::get_Path);
    long_getter!(IRegisteredTask::get_LastTaskResult);


    pub fn get_last_runtime(&mut self) -> Result<String, HResult> {
        unsafe {
            let mut date: DATE = 0 as DATE;
            com_call!(self.0, IRegisteredTask::get_LastRunTime(&mut date))?;
            Ok(date_to_datetime(date))

        }
    }

    pub fn get_next_runtime(&mut self) -> Result<String, HResult> {
        unsafe {
            let mut date: DATE = 0 as DATE;
            com_call!(self.0, IRegisteredTask::get_NextRunTime(&mut date))?;
            Ok(date_to_datetime(date))
        }
    }

    pub fn set_sd(&mut self, sddl: &BString) -> Result<(), HResult> {
        unsafe {
            com_call!(
                self.0,
                IRegisteredTask::SetSecurityDescriptor(
                    sddl.as_raw_ptr(),
                    0, // flags (none)
                )
            )?;
        }
        Ok(())
    }
    // #[allow(non_snake_case)]
    // pub fn com_getter(&mut self) -> Result<String, HResult> {
    //     unsafe {
    //         let mut b = ptr::null_mut();
    //         let r = com_call!(self.0, IRegisteredTask::get_Name(&mut b))?;
    //         let s= OsString::from_wide(BString::from_raw(b).ok_or_else(|| HResult::new(r))?.as_ref());
    //         Ok(s.to_string_lossy().to_string())
    //     }
    // }
    #[allow(non_snake_case)]
    pub fn get_state_string(&mut self) -> Result<String, HResult> {
        unsafe {
            let mut b = 0;
            com_call!(self.0, IRegisteredTask::get_State(&mut b))?;
            let msg = match b  {
                taskschd::TASK_STATE_DISABLED => "disabled",
                taskschd::TASK_STATE_QUEUED => "queue",
                taskschd::TASK_STATE_READY => "ready",
                taskschd::TASK_STATE_RUNNING => "running",
                _ => "unknown",
            };
            Ok(msg.to_string())
        }
    }
    pub fn get_definition(&mut self) -> Result<TaskDefinition, HResult> {
        unsafe { com_call_getter!(|tc| self.0, IRegisteredTask::get_Definition(tc)) }
            .map(TaskDefinition)
    }

    pub fn run(&self) -> Result<(), HResult> {
        self.run_impl(Option::<&OsStr>::None)?;
        Ok(())
    }

    fn run_impl(&self, param: Option<impl AsRef<OsStr>>) -> Result<ComRef<IRunningTask>, HResult> {
        // Running with parameters isn't currently exposed.
        // param can also be an array of strings, but that is not supported here
        let param = if let Some(p) = param {
            Some(try_to_bstring!(p)?)
        } else {
            None
        };

        unsafe {
            com_call_getter!(
                |rt| self.0,
                IRegisteredTask::Run(param.as_ref().as_raw_variant(), rt)
            )
        }
    }
}