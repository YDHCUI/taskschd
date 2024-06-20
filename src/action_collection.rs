use comedy::com::ComRef;
use winapi::um::taskschd::{IActionCollection};
use crate::long_getter;

pub struct ActionCollection(pub ComRef<IActionCollection>);

impl ActionCollection {
    long_getter!(IActionCollection::get_Count);

}