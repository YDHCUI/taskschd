use comedy::com::ComRef;
use winapi::um::taskschd::IRegistrationInfo;
use crate::{string_getter, bstring_putter};

pub struct RegistrationInfo(pub ComRef<IRegistrationInfo>);

impl RegistrationInfo {
    bstring_putter!(IRegistrationInfo::put_Author);
    bstring_putter!(IRegistrationInfo::put_Description);
    string_getter!(IRegistrationInfo::get_Author);
    string_getter!(IRegistrationInfo::get_Description);

}