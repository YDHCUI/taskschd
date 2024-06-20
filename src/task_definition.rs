use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::ptr;
use comedy::com::ComRef;
use comedy::{com_call, com_call_getter, HResult};
use winapi::shared::ntdef::LONG;
use winapi::um::taskschd;
use winapi::um::taskschd::{IAction, ITaskDefinition, ITaskFolder, ITrigger, ITriggerCollection};
use crate::action_collection::ActionCollection;
use crate::trigger_boot::BootTrigger;
use crate::trigger_daily::DailyTrigger;
use crate::exec_action::ExecAction;
use crate::ole_utils::{BString, empty_variant, OptionBstringExt};
use crate::principal::Principal;
use crate::registered_task::RegisteredTask;
use crate::registration_info::RegistrationInfo;
use crate::task_folder::TaskFolder;
use crate::task_settings::TaskSettings;
use crate::trigger_event::EventTrigger;
use crate::trigger_logon::LogonTrigger;
use crate::trigger_monthly::MonthlyTrigger;
use crate::trigger_time::TimeTrigger;
use crate::trigger_weekly::WeeklyTrigger;

pub struct TaskDefinition(pub ComRef<ITaskDefinition>);

impl TaskDefinition {
    pub fn get_settings(&mut self) -> Result<TaskSettings, HResult> {
        unsafe { com_call_getter!(|s| self.0, ITaskDefinition::get_Settings(s)) }.map(TaskSettings)
    }

    pub fn get_registration_info(&mut self) -> Result<RegistrationInfo, HResult> {
        unsafe { com_call_getter!(|ri| self.0, ITaskDefinition::get_RegistrationInfo(ri)) }
            .map(RegistrationInfo)
    }

    pub fn get_principal(&mut self) -> Result<Principal, HResult> {
        unsafe { com_call_getter!(|ri| self.0, ITaskDefinition::get_Principal(ri)) }
            .map(Principal)
    }

    unsafe fn add_action<T: winapi::Interface>(
        &mut self,
        action_type: taskschd::TASK_ACTION_TYPE,
    ) -> Result<ComRef<T>, HResult> {
        use self::taskschd::IActionCollection;

        let actions = com_call_getter!(|ac| self.0, ITaskDefinition::get_Actions(ac))?;
        let action = com_call_getter!(|a| actions, IActionCollection::Create(action_type, a))?;
        action.cast()
    }

    pub fn get_exec_actions(&mut self) -> Result<Vec<String>, HResult> {
        use self::taskschd::IActionCollection;
        let mut exec_actions = Vec::<String>::new();
        unsafe {
            let mut actions = com_call_getter!(|ac| self.0, ITaskDefinition::get_Actions(ac)).map(ActionCollection)?;
            let count = actions.get_Count()?;
            for i in 1..=count {
                let action = com_call_getter!(|a| actions.0, IActionCollection::get_Item(i, a))?;
                let mut action_type = 0;
                com_call!(action, IAction::get_Type(&mut action_type))?;
                if action_type == taskschd::TASK_ACTION_EXEC {
                    let mut action_impl = ExecAction(action.cast()?);
                    let args = action_impl.get_Arguments().unwrap_or_default();
                    let msg = format!("{} {}", action_impl.get_Path()?, args);
                    exec_actions.push(msg)

                }
            }
            Ok(exec_actions)

        }
    }

    pub fn add_exec_action(&mut self) -> Result<ExecAction, HResult> {
        unsafe { self.add_action(taskschd::TASK_ACTION_EXEC) }.map(ExecAction)
    }

    pub unsafe fn add_trigger<T: winapi::Interface>(
        &mut self,
        trigger_type: taskschd::TASK_TRIGGER_TYPE2,
    ) -> Result<ComRef<T>, HResult> {
        let triggers = com_call_getter!(|tc| self.0, ITaskDefinition::get_Triggers(tc))?;
        let trigger = com_call_getter!(|t| triggers, ITriggerCollection::Create(trigger_type, t))?;
        trigger.cast()
    }

    pub fn add_daily_trigger(&mut self) -> Result<DailyTrigger, HResult> {
        unsafe { self.add_trigger(taskschd::TASK_TRIGGER_DAILY) }.map(DailyTrigger)
    }

    pub fn add_boot_trigger(&mut self) -> Result<BootTrigger, HResult> {
        unsafe { self.add_trigger(taskschd::TASK_TRIGGER_BOOT) }.map(BootTrigger)
    }

    pub fn get_all_triggers(&mut self) -> Result<Vec<String>, HResult> {
        let mut found_triggers = Vec::new();

        unsafe {
            let triggers = com_call_getter!(|tc| self.0, ITaskDefinition::get_Triggers(tc))?;
            let mut count = 0;
            com_call!(triggers, ITriggerCollection::get_Count(&mut count))?;

            // Item indexes start at 1
            for i in 1..=count {
                let trigger = com_call_getter!(|t| triggers, ITriggerCollection::get_Item(i, t))?;

                let mut trigger_type = 0;
                com_call!(trigger, ITrigger::get_Type(&mut trigger_type))?;
                // println!("trigger {}", trigger_type);
                let msg = match trigger_type {
                    taskschd::TASK_TRIGGER_EVENT => {
                        let mut trigger_impl = EventTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    }
                    taskschd::TASK_TRIGGER_TIME => {
                        let mut trigger_impl = TimeTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    },
                    taskschd::TASK_TRIGGER_DAILY=> {
                        let mut trigger_impl = DailyTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    },
                    taskschd::TASK_TRIGGER_WEEKLY=> {
                        let mut trigger_impl = WeeklyTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    },
                    taskschd::TASK_TRIGGER_MONTHLY => {
                        let mut trigger_impl = MonthlyTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    }
                    taskschd::TASK_TRIGGER_BOOT => {
                        let mut trigger_impl = BootTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    },
                    taskschd::TASK_TRIGGER_LOGON => {
                        let mut trigger_impl = LogonTrigger(trigger.cast()?);
                        trigger_impl.to_string()
                    },
                    _ => { format!("unknown type:{}",trigger_type)}
                };
                found_triggers.push(msg);
            }
        }

        Ok(found_triggers)
    }
    pub fn get_daily_triggers(&mut self) -> Result<Vec<DailyTrigger>, HResult> {
        let mut found_triggers = Vec::new();

        unsafe {
            let triggers = com_call_getter!(|tc| self.0, ITaskDefinition::get_Triggers(tc))?;
            let mut count = 0;
            com_call!(triggers, ITriggerCollection::get_Count(&mut count))?;

            // Item indexes start at 1
            for i in 1..=count {
                let trigger = com_call_getter!(|t| triggers, ITriggerCollection::get_Item(i, t))?;

                let mut trigger_type = 0;
                com_call!(trigger, ITrigger::get_Type(&mut trigger_type))?;
                println!("trigger type: {}", trigger_type);

                if trigger_type == taskschd::TASK_TRIGGER_DAILY {
                    found_triggers.push(DailyTrigger(trigger.cast()?))
                }
            }
        }

        Ok(found_triggers)
    }

    pub fn create(
        &mut self,
        folder: &mut TaskFolder,
        task_name: &BString,
        service_account: Option<&BString>,
    ) -> Result<RegisteredTask, HResult> {
        self.register_impl(folder, task_name, service_account, taskschd::TASK_CREATE)
    }

    fn register_impl(
        &mut self,
        folder: &mut TaskFolder,
        task_name: &BString,
        service_account: Option<&BString>,
        creation_flags: taskschd::TASK_CREATION,
    ) -> Result<RegisteredTask, HResult> {
        let task_definition = self.0.as_raw_ptr();

        let password = empty_variant();

        let logon_type = if service_account.is_some() {
            taskschd::TASK_LOGON_SERVICE_ACCOUNT
        } else {
            taskschd::TASK_LOGON_INTERACTIVE_TOKEN
        };

        let sddl = empty_variant();

        let registered_task = unsafe {
            com_call_getter!(
                |rt| folder.0,
                ITaskFolder::RegisterTaskDefinition(
                    task_name.as_raw_ptr(),
                    task_definition,
                    creation_flags as LONG,
                    service_account.as_raw_variant(),
                    password,
                    logon_type,
                    sddl,
                    rt,
                )
            )?
        };

        Ok(RegisteredTask(registered_task))
    }

    pub fn get_xml(task_definition: &ComRef<ITaskDefinition>) -> Result<OsString, String> {
        unsafe {
            let mut xml = ptr::null_mut();
            com_call!(task_definition, ITaskDefinition::get_XmlText(&mut xml))
                .map_err(|e| format!("{}", e))?;

            Ok(OsString::from_wide(
                BString::from_raw(xml)
                    .ok_or_else(|| "null xml".to_string())?
                    .as_ref(),
            ))
        }
    }
}