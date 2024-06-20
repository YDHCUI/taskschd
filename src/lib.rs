/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/. */

//! A partial type-safe interface for Windows Task Scheduler 2.0
//!
//! This provides structs thinly wrapping the taskschd interfaces, with methods implemented as
//! they've been needed for the update agent.
//!
//! If it turns out that much more flexibility is needed in task definitions, it may be worth
//! generating an XML string and using `ITaskFolder::RegisterTask` or
//! `ITaskDefinition::put_XmlText`, rather than adding more and more boilerplate here.
//!
//! See https://docs.microsoft.com/windows/win32/taskschd/task-scheduler-start-page for
//! Microsoft's documentation.



pub mod ole_utils;
pub mod com_macro;
pub mod task_settings;
pub mod task_definition;
pub mod registration_info;
pub mod trigger_daily;
pub mod exec_action;
pub mod registered_task;
pub mod task_folder;
pub mod task_service;
pub mod trigger_boot;
pub mod trigger_time;
pub mod repetition_pattern;
pub mod trigger_weekly;
pub mod trigger_monthly;
pub mod trigger_event;
pub mod trigger_logon;
mod action_collection;
mod principal;
mod task_folder_collection;
