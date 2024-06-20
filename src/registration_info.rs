use comedy::com::ComRef;
use winapi::um::taskschd::IRegistrationInfo;
use crate::{bstring_getter, bstring_putter};

pub struct RegistrationInfo(pub ComRef<IRegistrationInfo>);

impl RegistrationInfo {
    bstring_putter!(IRegistrationInfo::put_Author);
    bstring_putter!(IRegistrationInfo::put_Description);
    bstring_getter!(IRegistrationInfo::get_Author);
    bstring_getter!(IRegistrationInfo::get_Description);

}