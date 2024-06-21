/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

use std::convert::TryInto;
use std::ffi::OsStr;
use std::mem;
use std::os::windows::ffi::OsStrExt;
use std::ptr::NonNull;
use std::slice;
use chrono::NaiveDateTime;

use winapi::shared::{winerror, wtypes};
use winapi::um::{oaidl, oleauto, taskschd};

use comedy::{HResult, Win32Error};
use failure::Fail;
use winapi::shared::winerror::{ERROR_ALREADY_EXISTS, ERROR_FILE_NOT_FOUND};
use winapi::um::oaidl::VARIANT;
use winapi::um::oleauto::VariantInit;


#[derive(Debug)]
pub struct BString(NonNull<u16>);

impl BString {
    pub fn from_slice(v: impl AsRef<[u16]>) -> Result<BString, HResult> {
        let v = v.as_ref();
        let real_len = v.len();
        let len = real_len
            .try_into()
            .map_err(|_| HResult::new(winerror::E_OUTOFMEMORY))?;
        let bs = unsafe { oleauto::SysAllocStringLen(v.as_ptr(), len) };

        Ok(Self(NonNull::new(bs).ok_or_else(|| {
            HResult::new(winerror::E_OUTOFMEMORY).function("SysAllocStringLen")
        })?))
    }

    pub fn from_os_str(s: impl AsRef<OsStr>) -> Result<BString, HResult> {
        BString::from_slice(s.as_ref().encode_wide().collect::<Vec<_>>().as_slice())
    }

    pub unsafe fn from_raw(p: *mut u16) -> Option<Self> {
        Some(Self(NonNull::new(p)?))
    }
    pub fn as_raw_ptr(&self) -> *mut u16 {
        self.0.as_ptr()
    }

    pub fn as_raw_variant(&self) -> oaidl::VARIANT {
        unsafe {
            let mut v: oaidl::VARIANT = mem::zeroed();
            {
                let tv = v.n1.n2_mut();
                *tv.n3.bstrVal_mut() = self.as_raw_ptr();
                tv.vt = wtypes::VT_BSTR as wtypes::VARTYPE;
            }

            v
        }
    }
    // 添加一个方法来返回 BString 的长度
    pub fn len(&self) -> u32 {
        unsafe { oleauto::SysStringLen(self.0.as_ptr()) }
    }

    // 将 BString 转换回 Rust String
    pub fn to_string(&self) -> String {
        let len = self.len() as usize;
        let slice = unsafe { slice::from_raw_parts(self.0.as_ptr(), len) };
        String::from_utf16_lossy(slice)
    }

}

impl Drop for BString {
    fn drop(&mut self) {
        unsafe { oleauto::SysFreeString(self.0.as_ptr()) }
    }
}

impl AsRef<[u16]> for BString {
    fn as_ref(&self) -> &[u16] {
        unsafe {
            let len = oleauto::SysStringLen(self.0.as_ptr());

            slice::from_raw_parts(self.0.as_ptr(), len as usize)
        }
    }
}

/// Try to convert, decorate `Err` with call site info
#[macro_export]
macro_rules! try_to_bstring {
    ($ex:expr) => {
        $crate::ole_utils::BString::from_os_str($ex).map_err(|e| e.file_line(file!(), line!()))
    };
}

pub fn empty_variant() -> oaidl::VARIANT {
    unsafe {
        let mut v: oaidl::VARIANT = mem::zeroed();
        {
            let tv = v.n1.n2_mut();
            tv.vt = wtypes::VT_EMPTY as wtypes::VARTYPE;
        }

        v
    }
}

pub trait OptionBstringExt {
    fn as_raw_variant(&self) -> oaidl::VARIANT;
}

/// Shorthand for unwrapping, returns `BString::as_raw_variant()` or `empty_variant()`
impl OptionBstringExt for Option<&BString> {
    fn as_raw_variant(&self) -> oaidl::VARIANT {
        self.map(|bs| bs.as_raw_variant())
            .unwrap_or_else(empty_variant)
    }
}

// Note: A `VARIANT_BOOL` is not a `VARIANT`, rather it would go into a `VARIANT` of type
// `VT_BOOL`. Some APIs use it directly.
pub trait IntoVariantBool {
    fn into_variant_bool(self) -> wtypes::VARIANT_BOOL;
}

impl IntoVariantBool for bool {
    fn into_variant_bool(self) -> wtypes::VARIANT_BOOL {
        if self {
            wtypes::VARIANT_TRUE
        } else {
            wtypes::VARIANT_FALSE
        }
    }
}


pub trait IntoVariantI32 {
    fn into_variant_i32(self) -> VARIANT;
}
impl IntoVariantI32 for i32 {
    fn into_variant_i32(self) -> VARIANT {
        unsafe {
            // 初始化 VARIANT
            let mut var = mem::zeroed::<VARIANT>();
            VariantInit(&mut var);
            // 设置 VARIANT 为整数类型 (VT_I4) 并赋值
            (*var.n1.n2_mut()).vt = wtypes::VT_I4 as u16; // 设置 VARTYPE
            *(*var.n1.n2_mut()).n3.lVal_mut() = self;
            var
        }
    }
}


pub fn hr_is_not_found(hr: &HResult) -> bool {
    hr.code() == HResult::from(Win32Error::new(ERROR_FILE_NOT_FOUND)).code()
}

pub fn hr_is_already_exists(hr: &HResult) -> bool {
    hr.code() == HResult::from(Win32Error::new(ERROR_ALREADY_EXISTS)).code()
}

pub fn date_to_datetime(date: wtypes::DATE) -> String {
    const OFFSET_DAYS: i32 = 25569; // 从 "1970-01-01" 到 "1899-12-30" 的天数
    const SECONDS_PER_DAY: i64 = 86_400; // 一天的秒数

    // 将date转换为Unix时间戳（从 "1970-01-01" 开始）
    let timestamp = (date - OFFSET_DAYS as f64) * SECONDS_PER_DAY as f64;

    // 创建一个 'NaiveDateTime' 的实例
    match NaiveDateTime::from_timestamp_opt(timestamp as i64, 0)  {
        None => {"".to_string()}
        Some(val) => {val.to_string()}
    }
}



#[derive(Clone, Debug, Fail)]
pub enum ConnectTaskServiceError {
    #[fail(display = "{}", _0)]
    CreateInstanceFailed(#[fail(cause)] HResult),
    #[fail(display = "Access is denied to connect to the Task Scheduler service")]
    AccessDenied(#[fail(cause)] HResult),
    #[fail(display = "The Task Scheduler service is not running")]
    ServiceNotRunning(#[fail(cause)] HResult),
    #[fail(display = "{}", _0)]
    ConnectFailed(#[fail(cause)] HResult),
}





#[derive(Clone, Copy, Debug)]
#[repr(u32)]
pub enum InstancesPolicy {
    Parallel = taskschd::TASK_INSTANCES_PARALLEL,
    Queue = taskschd::TASK_INSTANCES_QUEUE,
    IgnoreNew = taskschd::TASK_INSTANCES_IGNORE_NEW,
    StopExisting = taskschd::TASK_INSTANCES_STOP_EXISTING,
}