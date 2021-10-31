
use std::{thread, time::Duration};
use windows::{
    Win32::Foundation::PWSTR,
    Win32::UI::WindowsAndMessaging::{
        GetForegroundWindow,
        GetWindowTextLengthW,
        GetWindowTextW,
    },
};
use anyhow::{
    Result,
    anyhow,
};

fn main() {
    loop {
        unsafe {
            match get_active_window_title() {
                Ok(title) => {
                    println!("Window title: {}", title);
                }
                Err(error) => {
                    println!("Error: {}", error);
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