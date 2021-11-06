use core::ops::DerefMut;
use std::{
    thread,
    time::Duration
};
use windows::Win32::{
    Foundation::*,
    System::Com::*,
    System::{
        OleAutomation::*,
        PropertiesSystem::*,
    },
    UI::{
        WindowsAndMessaging::*,
        Accessibility::*,
    },
};
use anyhow::{
    Result,
    anyhow,
};
use widestring::U16CStr;

// error: process didn't exit successfully: `target\debug\window-title-getter.exe` (exit code: 0xc0000374, STATUS_HEAP_CORRUPTION)
fn main() {
    loop {
        unsafe {
            match get_active_window_title() {
                Ok(title) => {
                    println!("Window title: {}", title);
                }
                Err(error) => {
                    println!("Error getting window title: {}", error);
                }
            }
            match get_active_browser_url() {
                Ok(url) => {
                    println!("Browser URL: {}", url);
                }
                Err(error) => {
                    println!("Error getting browser URL: {}", error);
                }
            }
        }
        thread::sleep(Duration::from_secs(1));
    }
}

unsafe fn get_active_window_title() -> Result<String> {
    let foreground_window_hwnd = GetForegroundWindow();
    if foreground_window_hwnd.0 > 0 {
        let title_length = GetWindowTextLengthW(foreground_window_hwnd);
        if title_length > 0 {
            let mut title: Vec<u16> = vec![0; title_length as usize];
            let pwstr_title = PWSTR(title.as_mut_ptr());
            let received_title_length = GetWindowTextW(foreground_window_hwnd, pwstr_title, title_length + 1);
            if received_title_length > 0 {
                let title = String::from_utf16(&title)?;
                Ok(title)
            }
            else {
                Err(anyhow!("Received window title length is invalid: {}", received_title_length))
            }
        }
        else {
            Err(anyhow!("Window title length is invalid: {}", title_length))
        }
    }
    else {
        Err(anyhow!("Foreground window hwnd is invalid: {}", foreground_window_hwnd.0))
    }
}

unsafe fn get_active_browser_url() -> Result<String> {
    let foreground_window_hwnd = GetForegroundWindow();
    CoInitializeEx(std::ptr::null(), COINIT_MULTITHREADED)?;
    let ui_automation: IUIAutomation = CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?;
    let ui_element = ui_automation.ElementFromHandle(foreground_window_hwnd)?;
    get_chrome_url(ui_automation, ui_element)
}

unsafe fn get_chrome_url(ui_automation: IUIAutomation, ui_element: IUIAutomationElement) -> Result<String> {
    let search_condition = create_search_condition_for_url_editor(ui_automation)?;
    let edit_element = ui_element.FindFirst(TreeScope_Descendants, search_condition)?;
    get_value_from_ui_element(edit_element)
}

unsafe fn create_search_condition_for_url_editor(ui_automation: IUIAutomation) -> Result<IUIAutomationCondition> {
    const CHROME_CONTROL_ID: i32 = 0xC354;
    let mut variant = VARIANT::default();
    variant.Anonymous.Anonymous.deref_mut().vt = VT_I4.0 as u16;
    variant.Anonymous.Anonymous.deref_mut().Anonymous.lVal = CHROME_CONTROL_ID;

    let search_condition = ui_automation.CreatePropertyCondition(UIA_ControlTypePropertyId, variant)?;
    Ok(search_condition)
}

unsafe fn get_value_from_ui_element(ui_element: IUIAutomationElement) -> Result<String> {
    let value = ui_element.GetCurrentPropertyValue(UIA_ValueValuePropertyId)?;
    let title = VariantToStringAlloc(&value)?;
    let title = U16CStr::from_ptr_str(title.0).to_string()?;
    Ok(title)
}