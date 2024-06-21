
/// put a bool, converting to `VARIANT_BOOL`
#[macro_export]
macro_rules! bool_putter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self, v: bool) -> Result<(), comedy::HResult> {
            use crate::ole_utils::IntoVariantBool;
            let v = v.into_variant_bool();
            unsafe {
                comedy::com_call!(self.0, $interface::$method(v))?;
            }
            Ok(())
        }
    };
}

#[macro_export]
macro_rules! bool_getter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self) -> Result<bool, comedy::HResult> {
            let mut v =  0;
            unsafe {
                comedy::com_call!(self.0, $interface::$method(&mut v))?;
            }
            if v == 0 {
                Ok(false)
            } else {
                Ok(true)
            }
        }
    };
}


/// put a value that is already available as a `BString`
#[macro_export]
macro_rules! bstring_putter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self, v: &crate::ole_utils::BString) -> Result<(), comedy::HResult> {
            unsafe {
                comedy::com_call!(self.0, $interface::$method(v.as_raw_ptr()))?;
            }
            Ok(())
        }
    };
}

#[macro_export]
macro_rules! string_getter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self) -> Result<String, comedy::HResult> {
            use std::os::windows::ffi::OsStringExt;
            unsafe {
                let mut b = std::ptr::null_mut();
                let r = comedy::com_call!(self.0, $interface::$method(&mut b))?;
                let s= std::ffi::OsString::from_wide(crate::ole_utils::BString::from_raw(b).ok_or_else(|| comedy::HResult::new(r))?.as_ref());
                Ok(s.to_string_lossy().to_string())
            }
        }
    };
}

#[macro_export]
macro_rules! short_getter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self) -> Result<winapi::shared::ntdef::SHORT, comedy::HResult> {
            unsafe {
                let mut i = 0;
                comedy::com_call!(self.0, $interface::$method(&mut i))?;
                Ok(i)
            }
        }
    };
}

#[macro_export]
macro_rules! long_getter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self) -> Result<winapi::shared::ntdef::LONG, comedy::HResult> {
            unsafe {
                let mut i = 0;
                comedy::com_call!(self.0, $interface::$method(&mut i))?;
                Ok(i)
            }
        }
    };
}
#[macro_export]
macro_rules! get_repetition {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self) -> Result<String, comedy::HResult> {
            use winapi::um::taskschd::{ITrigger};

            unsafe {
                let mut repeat = std::ptr::null_mut();
                let r = comedy::com_call!(self.0, ITrigger::get_Repetition(&mut repeat))?;
                // 将 *mut IRepetitionPattern 转换为 NonNull<IRepetitionPattern>
                if let Some(nonnull) = std::ptr::NonNull::new(repeat) {
                    // 使用 NonNull<IRepetitionPattern> 构造 ComRef<IRepetitionPattern>
                    let com_ref = comedy::com::ComRef::from_raw(nonnull);
                    // 使用 ComRef<IRepetitionPattern> 构造 RepetitionPattern
                    let mut repetition_pattern = crate::repetition_pattern::RepetitionPattern(com_ref);
                    Ok(repetition_pattern.to_string())
                    // 现在你可以使用 repetition_pattern 了
                } else {
                    Err(comedy::HResult::new(r))
                }
            }
        }
    };
}
/// put a `chrono::DateTime` value
#[macro_export]
macro_rules! datetime_putter {
    ($interface:ident :: $method:ident) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self, v: chrono::DateTime<chrono::Utc>) -> Result<(), comedy::HResult> {
            let v = crate::try_to_bstring!(v.to_rfc3339_opts(chrono::SecondsFormat::Secs, true))?;
            unsafe {
                comedy::com_call!(self.0, $interface::$method(v.as_raw_ptr()))?;
            }
            Ok(())
        }
    };
}


/// put a value of type `$ty`, which implements `AsRef<OsStr>`
#[macro_export]
macro_rules! to_os_str_putter {
    ($interface:ident :: $method:ident, $ty:ty) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self, v: $ty) -> Result<(), comedy::HResult> {
            let v = crate::try_to_bstring!(v)?;
            unsafe {
                comedy::com_call!(self.0, $interface::$method(v.as_raw_ptr()))?;
            }
            Ok(())
        }
    };
}

/// put a value of type `$ty`, which implements `ToString`
#[macro_export]
macro_rules! to_string_putter {
    ($interface:ident :: $method:ident, $ty:ty) => {
        #[allow(non_snake_case)]
        pub fn $method(&mut self, v: $ty) -> Result<(), comedy::HResult> {
            let v = crate::try_to_bstring!(v.to_string())?;
            unsafe {
                comedy::com_call!(self.0, $interface::$method(v.as_raw_ptr()))?;
            }
            Ok(())
        }
    };
}