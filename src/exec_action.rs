use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use comedy::com::ComRef;
use comedy::{com_call, HResult, Win32Error};
use winapi::shared::winerror::ERROR_BAD_ARGUMENTS;
use winapi::um::taskschd::IExecAction;
use crate::ole_utils::BString;
use crate::{string_getter, bstring_putter, to_os_str_putter, try_to_bstring};

pub struct ExecAction(pub ComRef<IExecAction>);

impl ExecAction {
    string_getter!(IExecAction::get_Path);
    string_getter!(IExecAction::get_Arguments);


    to_os_str_putter!(IExecAction::put_Path, &Path);
    to_os_str_putter!(IExecAction::put_WorkingDirectory, &Path);

    pub fn put_path(&mut self, path: &str) -> Result<(), HResult> {
        unsafe  {
            com_call!(self.0, IExecAction::put_Path(try_to_bstring!(path)?.as_raw_ptr()))?;
        }
        Ok(())
    }
    #[allow(non_snake_case)]
    pub fn put_Arguments(&mut self, args: &[OsString]) -> Result<(), HResult> {
        // based on `make_command_line()` from libstd
        // https://github.com/rust-lang/rust/blob/37ff5d388f8c004ca248adb635f1cc84d347eda0/src/libstd/sys/windows/process.rs#L457

        let mut s = Vec::new();

        fn append_arg(cmd: &mut Vec<u16>, arg: &OsStr) -> Result<(), HResult> {
            cmd.push('"' as u16);

            let mut backslashes: usize = 0;
            for x in arg.encode_wide() {
                if x == 0 {
                    return Err(HResult::from(Win32Error::new(ERROR_BAD_ARGUMENTS))
                        .file_line(file!(), line!()));
                }

                if x == '\\' as u16 {
                    backslashes += 1;
                } else {
                    if x == '"' as u16 {
                        // Add n+1 backslashes for a total of 2n+1 before internal '"'.
                        cmd.extend((0..=backslashes).map(|_| '\\' as u16));
                    }
                    backslashes = 0;
                }
                cmd.push(x);
            }

            // Add n backslashes for a total of 2n before ending '"'.
            cmd.extend((0..backslashes).map(|_| '\\' as u16));
            cmd.push('"' as u16);

            Ok(())
        }

        for arg in args {
            if !s.is_empty() {
                s.push(' ' as u16);
            }

            // always quote args
            append_arg(&mut s, arg.as_ref())?;
        }

        let args = BString::from_slice(s).map_err(|e| e.file_line(file!(), line!()))?;

        unsafe {
            com_call!(self.0, IExecAction::put_Arguments(args.as_raw_ptr()))?;
        }
        Ok(())
    }
}
