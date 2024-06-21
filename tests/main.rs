extern crate taskschd;

use std::ffi::{OsStr, OsString};
use std::path::Path;
use taskschd::ole_utils::hr_is_not_found;
use taskschd::task_service::TaskService;

use taskschd::try_to_bstring;


#[test]
fn get_tasks_test() -> Result<(), failure::Error>{
    let mut service = TaskService::connect_local()?;
    let mut tasks = service.get_all_tasks()?;
    for mut task in tasks {
        println!("task path: {}, state: {}, last_runtime: {}", task.path, task.state, task.user_id);
    }
    Ok(())
}

#[test]
fn stop_tasks_test() -> Result<(), failure::Error>{
    let mut service = TaskService::connect_local()?;
    let mut folder = service.get_folder(&try_to_bstring!(r#"\Microsoft\Windows\WindowsUpdate"#)?)?;
    let mut task = folder.get_task(&try_to_bstring!("Update")?)?;
    // task.put_Enabled(false)?;
    // task.put_Enabled(true)?;
    // task.run()?;
    // let mut action = task.get_definition()?.add_exec_action()?;
    // let a = OsString::from("--help");
    //  action.put_Arguments(&[a])?;

    let mut task_def = task.get_definition()?;

    // let mut task =  task_def.update(&mut folder, &try_to_bstring!("Update")?, None)?;
    let actions = task_def.get_exec_actions()?;
    for mut action in actions {
        println!("before: {}", action.get_Path()?);

        action.put_path("c:\\windows\\system32\\calc1.exe")?;
        println!("after: {}", action.get_Path()?);
    }
    task_def.update(&mut folder, &try_to_bstring!("Update")?)?;
    Ok(())
}
#[test]
fn update_task_user_id_test() -> Result<(), failure::Error>{
    let mut service = TaskService::connect_local()?;
    let mut folder = service.get_folder(&try_to_bstring!(r#"\Microsoft\Windows\WindowsUpdate"#)?)?;
    let mut task = folder.get_task(&try_to_bstring!("Update")?)?;

    let mut task_def = task.get_definition()?;
    task_def.get_principal()?.put_UserId("Administrator")?;

    task_def.update(&mut folder, &try_to_bstring!("Update")?)?;
    Ok(())
}
#[test]
fn register() -> Result<(), failure::Error>{
    
    let task_name = try_to_bstring!("name")?;
    let task_folder = try_to_bstring!("\\")?;
    let task_exe = Path::new("C:\\Windows\\System32\\cmd.exe");
    let task_args = vec![OsString::from("/c ping 127.0.0.1 -t")];

    let mut service = TaskService::connect_local()?;

    // Get or create the folder
    let mut folder = service.get_folder(&task_folder).or_else(|e| {
        if hr_is_not_found(&e) {
            service
                .get_root_folder()
                .and_then(|mut root| root.create_folder(&task_folder))
        } else {
            Err(e)
        }
    })?;


    let start_time = folder
            .get_task(&task_name)
            .ok()
            .and_then(|mut task| task.get_definition().ok())
            .and_then(|mut def| def.get_daily_triggers().ok())
            .and_then(|mut triggers| {
                // Currently we are only using 1 daily trigger.
                triggers
                    .get_mut(0)
                    .and_then(|trigger| trigger.get_StartBoundary().ok())
            });

    let _ = folder.delete_task(&task_name);

    let mut task_def = service.new_task_definition()?;

    {


        let mut action = task_def.add_exec_action()?;
        action.put_Path(task_exe)?;
        //action.put_Arguments(task_args.as_slice())?;
        // TODO working directory?
    }

    {
        let mut settings = task_def.get_settings()?;
        settings.put_DisallowStartIfOnBatteries(false)?;
        settings.put_StopIfGoingOnBatteries(false)?;
        settings.put_StartWhenAvailable(true)?;
        settings.put_ExecutionTimeLimit(chrono::Duration::minutes(5))?;
    }

    {
        let mut info = task_def.get_registration_info()?;
        info.put_Author(&try_to_bstring!("author")?)?;
        info.put_Description(&try_to_bstring!("task test")?)?;
    }

    // A daily trigger starting 5 minutes ago.
    {
        let mut daily_trigger = task_def.add_daily_trigger()?;
        if let Some(ref start_time) = start_time {
            let s = try_to_bstring!(start_time)?;
            daily_trigger.put_StartBoundary_BString(&s)?;
        } else {
            daily_trigger.put_StartBoundary(chrono::Utc::now() - chrono::Duration::minutes(5))?;
        }
        daily_trigger.put_DaysInterval(1)?;
        // TODO: 12-hourly trigger? logon trigger?
    }

    let service_account = Some(try_to_bstring!("NT AUTHORITY\\LocalService")?);

    let mut registered_task = task_def.create(&mut folder, &task_name, service_account.as_ref())?;

    let sddl = try_to_bstring!(concat!(
            "D:(",   // DACL
            "A;",    // ace_type = Allow
            ";",     // ace_flags = none
            "GRGX;", // rights = Generic Read, Generic Execute
            ";;",    // object_guid, inherit_object_guid = none
            "BU)"    // account_sid = Built-in users
        ))?;

    registered_task.set_sd(&sddl)?;
    
    Ok(())
}

#[test]
fn unregister() -> Result<(), failure::Error> {
    let task_name = try_to_bstring!("name")?;
    let task_folder = try_to_bstring!("\\")?;

    let mut service = TaskService::connect_local()?;
    let maybe_folder = service.get_folder(&task_folder);
    let mut folder = match maybe_folder {
        Err(e) => {
            if hr_is_not_found(&e) {
                return Ok(());
            } else {
                return Err(e.into());
            }
        }
        Ok(folder) => folder,
    };

    folder.delete_task(&task_name).or_else(|e| {
        if hr_is_not_found(&e) {
            Ok(())
        } else {
            // Other errors are fatal.
            Err(e)
        }
    })?;

    let count = folder.get_task_count(true).unwrap_or_else(|e| {
        1
    });

    if count == 0 {
        let result = service
            .get_root_folder()
            .and_then(|mut root| root.delete_folder(&task_folder));
    }

    Ok(())
}

#[test]
fn run_on_demand() -> Result<(), failure::Error> {
    let task_name = try_to_bstring!("name")?;
    let task_folder = try_to_bstring!("\\")?;

    let mut service = TaskService::connect_local()?;
    let task = service.get_folder(&task_folder)?.get_task(&task_name)?;

    task.run()?;

    Ok(())
}

