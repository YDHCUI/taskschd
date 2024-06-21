use comedy::com::{ComRef, create_instance_inproc_server, INIT_MTA};
use comedy::{com_call, com_call_getter, HResult};
use winapi::shared::winerror::{E_ACCESSDENIED, SCHED_E_SERVICE_NOT_RUNNING};
use winapi::um::taskschd;
use winapi::um::taskschd::ITaskService;
use crate::ole_utils::{BString, ConnectTaskServiceError, empty_variant};
use crate::task_definition::TaskDefinition;
use crate::task_folder::{TaskFolder, TaskInfo};
use crate::try_to_bstring;

pub struct TaskService(pub ComRef<ITaskService>);

impl TaskService {
    pub fn connect_local() -> Result<TaskService, ConnectTaskServiceError> {
        use self::ConnectTaskServiceError::*;

        INIT_MTA.with(|com| {
            let _com = match com {
                Err(e) => return Err(e.clone()),
                Ok(ref _com) => _com,
            };
            //do_com_stuff(com);
            Ok(())
        }).map_err(CreateInstanceFailed)?;

        let task_service = create_instance_inproc_server::<taskschd::TaskScheduler, ITaskService>()
            .map_err(CreateInstanceFailed)?;

        // Connect to local service with no credentials.
        unsafe {
            com_call!(
                task_service,
                ITaskService::Connect(
                    empty_variant(),
                    empty_variant(),
                    empty_variant(),
                    empty_variant()
                )
            )
        }
            .map_err(|hr| match hr.code() {
                E_ACCESSDENIED => AccessDenied(hr),
                SCHED_E_SERVICE_NOT_RUNNING => ServiceNotRunning(hr),
                _ => ConnectFailed(hr),
            })?;

        Ok(TaskService(task_service))
    }

    pub fn get_root_folder(&mut self) -> Result<TaskFolder, HResult> {
        self.get_folder(&try_to_bstring!("\\")?)
    }

    pub fn get_all_tasks(&mut self) -> Result<Vec<TaskInfo>, HResult> {
        let mut task_infos = Vec::<TaskInfo>::new();
        let mut root = self.get_root_folder()?;
        let mut folders = root.get_all_folders()?;
        folders.insert(0, root);
        for mut folder in folders {
            if let Ok(infos) = folder.get_tasks(true) {
                if infos.is_empty() {
                    continue
                }
                task_infos.extend(infos);
            }
        };
        Ok(task_infos)

    }

    pub fn get_folder(&mut self, path: &BString) -> Result<TaskFolder, HResult> {
        unsafe {
            com_call_getter!(
                |folder| self.0,
                ITaskService::GetFolder(path.as_raw_ptr(), folder)
            )
        }
            .map(TaskFolder)
    }

    pub fn new_task_definition(&mut self) -> Result<TaskDefinition, HResult> {
        unsafe {
            com_call_getter!(
                |task_def| self.0,
                ITaskService::NewTask(
                    0, // flags (reserved)
                    task_def,
                )
            )
        }
            .map(TaskDefinition)
    }
}
