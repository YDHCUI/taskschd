use comedy::com::ComRef;
use comedy::{com_call, HResult};
use winapi::um::taskschd::ITaskSettings;
use crate::{bool_putter, to_string_putter};
use crate::ole_utils::InstancesPolicy;

pub struct TaskSettings(pub ComRef<ITaskSettings>);

impl TaskSettings {
    bool_putter!(ITaskSettings::put_AllowDemandStart);
    bool_putter!(ITaskSettings::put_DisallowStartIfOnBatteries);
    to_string_putter!(ITaskSettings::put_ExecutionTimeLimit, chrono::Duration);
    bool_putter!(ITaskSettings::put_Hidden);

    #[allow(non_snake_case)]
    pub fn put_MultipleInstances(&mut self, v: InstancesPolicy) -> Result<(), HResult> {
        unsafe {
            com_call!(self.0, ITaskSettings::put_MultipleInstances(v as u32))?;
        }
        Ok(())
    }

    bool_putter!(ITaskSettings::put_RunOnlyIfIdle);
    bool_putter!(ITaskSettings::put_RunOnlyIfNetworkAvailable);
    bool_putter!(ITaskSettings::put_StartWhenAvailable);
    bool_putter!(ITaskSettings::put_StopIfGoingOnBatteries);
    bool_putter!(ITaskSettings::put_Enabled);
    bool_putter!(ITaskSettings::put_WakeToRun);
}