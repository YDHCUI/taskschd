use comedy::com::ComRef;
use comedy::{com_call, com_call_getter, HResult};
use comedy::error::{ErrorAndSource, HResultInner};
use winapi::shared::ntdef::LONG;
use winapi::um::taskschd;
use winapi::um::taskschd::{ITaskFolder, ITaskFolderCollection};
use crate::string_getter;
use crate::ole_utils::{BString, empty_variant, IntoVariantI32};
use crate::registered_task::RegisteredTask;
use crate::registration_info::RegistrationInfo;
use crate::task_definition::TaskDefinition;


pub struct TaskInfo {
    pub name: String,
    pub path: String,
    pub state: String,
    pub description: String,
    pub triggers: String,

    pub exec_actions: String,

    pub last_runtime: String,
    pub next_runtime: String,
    pub last_task_result: i32,

    pub author: String,
    pub principal: String,
    pub user_id: String,
    pub xml: String,

}

pub struct TaskFolder(pub(crate) ComRef<ITaskFolder>);

impl TaskFolder {
    string_getter!(ITaskFolder::get_Name);
    string_getter!(ITaskFolder::get_Path);

    pub fn get_task(&mut self, task_name: &BString) -> Result<RegisteredTask, HResult> {
        unsafe {
            com_call_getter!(
                |task| self.0,
                ITaskFolder::GetTask(task_name.as_raw_ptr(), task)
            )
        }
            .map(RegisteredTask)
    }

    pub fn get_all_folders(&mut self) -> Result<Vec<TaskFolder>, HResult>{
        let mut found_folders = Vec::<TaskFolder>::with_capacity(1);
        unsafe {
            let mut folders = std::ptr::null_mut();
            com_call!(self.0, ITaskFolder::GetFolders(0, &mut folders))?;
            let mut count = 0;
            com_call!(folders, ITaskFolderCollection::get_Count(&mut count))?;
            for i in 1..=count {
                let mut folder = com_call_getter!(|t| folders, ITaskFolderCollection::get_Item(i.into_variant_i32(), t)).map(TaskFolder)?;
                if let Ok(child_folders) = folder.get_all_folders() {
                    found_folders.extend(child_folders);
                }
                found_folders.insert(0, folder);
            }
        }
        Ok(found_folders)
    }

    pub fn get_task_count(&mut self, include_hidden: bool) -> Result<LONG, HResult> {
        use self::taskschd::IRegisteredTaskCollection;

        let flags = if include_hidden {
            taskschd::TASK_ENUM_HIDDEN
        } else {
            0
        };

        unsafe {
            let tasks = com_call_getter!(|t| self.0, ITaskFolder::GetTasks(flags as LONG, t))?;

            let mut count = 0;
            com_call!(tasks, IRegisteredTaskCollection::get_Count(&mut count))?;

            Ok(count)
        }
    }

    pub fn get_tasks(&mut self, include_hidden: bool) -> Result<Vec<TaskInfo>, HResult> {
        use self::taskschd::IRegisteredTaskCollection;

        let flags = if include_hidden {
            taskschd::TASK_ENUM_HIDDEN
        } else {
            0
        };
        let mut task_infos = Vec::<TaskInfo>::new();


        unsafe {
            let tasks = com_call_getter!(|t| self.0, ITaskFolder::GetTasks(flags as LONG, t))?;
            let mut count = 0;
            com_call!(tasks, IRegisteredTaskCollection::get_Count(&mut count))?;
            for i in 1..=count {
                let mut task = match com_call_getter!(|t| tasks, IRegisteredTaskCollection::get_Item(i.into_variant_i32(), t)).map(RegisteredTask) {
                    Ok(val) => {val}
                    Err(_) => {continue}
                };
                let mut definition = match task.get_definition() {
                    Ok(val) => {val}
                    Err(_) => {continue}
                };
                let mut registration_info = match definition.get_registration_info() {
                    Ok(val) => {val}
                    Err(_) => {continue}
                };
                let mut task_info = TaskInfo{
                    name: task.get_Name().unwrap_or_default(),
                    path: task.get_Path().unwrap_or_default(),
                    state: task.get_state_string().unwrap_or_default(),
                    author: registration_info.get_Author().unwrap_or_default(),
                    description: registration_info.get_Description().unwrap_or_default(),
                    principal: "".to_string(),
                    triggers: "".to_string(),
                    exec_actions: "".to_string(),
                    last_runtime: task.get_last_runtime().unwrap_or_default(),
                    next_runtime: task.get_next_runtime().unwrap_or_default(),
                    xml: "".to_string(),
                    last_task_result: 0,
                    user_id: "".to_string(),
                };

                if let Ok(result) = task.get_LastTaskResult(){
                    task_info.last_task_result = result
                }

                if let Ok(mut principal) = definition.get_principal() {
                    task_info.principal = principal.to_string();
                    task_info.user_id = principal.get_UserId().unwrap_or_default();
                }
                let xml = TaskDefinition::get_xml(&definition.0).unwrap_or_default();
                task_info.xml = xml.to_string_lossy().to_string();




                if let Ok(triggers) = definition.get_all_triggers() {
                    task_info.triggers = triggers.join("\n")
                }
                if let Ok(exec_actions) = definition.get_exec_actions_string() {
                    task_info.exec_actions = exec_actions.join("\n")
                }
                task_infos.push(task_info)
            }

            Ok(task_infos)
        }
    }

    pub fn create_folder(&mut self, path: &BString) -> Result<TaskFolder, HResult> {
        let sddl = empty_variant();
        unsafe {
            com_call_getter!(
                |folder| self.0,
                ITaskFolder::CreateFolder(path.as_raw_ptr(), sddl, folder)
            )
        }
            .map(TaskFolder)
    }

    pub fn delete_folder(&mut self, path: &BString) -> Result<(), HResult> {
        unsafe {
            com_call!(
                self.0,
                ITaskFolder::DeleteFolder(
                    path.as_raw_ptr(),
                    0, // flags (reserved)
                )
            )?;
        }

        Ok(())
    }

    pub fn delete_task(&mut self, task_name: &BString) -> Result<(), HResult> {
        unsafe {
            com_call!(
                self.0,
                ITaskFolder::DeleteTask(
                    task_name.as_raw_ptr(),
                    0, // flags (reserved)
                )
            )?;
        }

        Ok(())
    }
}

