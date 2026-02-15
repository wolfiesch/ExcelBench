use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

pub(super) fn looks_like_date_format(code: &str) -> bool {
    // Heuristic: date formats typically include year + day tokens.
    let lc = code.to_ascii_lowercase();
    lc.contains('y') && lc.contains('d')
}

pub(super) fn excel_serial_to_naive_datetime(serial: f64) -> Option<NaiveDateTime> {
    // Excel 1900 date system, with the standard 1900 leap-year bug adjustment.
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let mut f = serial;
    if f < 60.0 {
        f += 1.0;
    }
    let total_ms = (f * 86_400_000.0).round() as i64;
    epoch.checked_add_signed(Duration::milliseconds(total_ms))
}

pub(super) fn naive_datetime_to_excel_serial(dt: NaiveDateTime) -> Option<f64> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let delta = dt - epoch;
    let total_ms = delta.num_milliseconds();
    Some(total_ms as f64 / 86_400_000.0)
}

// ---------------------------------------------------------------------------
// Color helpers: ARGB <-> hex
// ---------------------------------------------------------------------------

/// Convert ARGB "FFRRGGBB" or "RRGGBB" to "#RRGGBB".
pub(super) fn argb_to_hex(argb: &str) -> String {
    let s = argb.trim();
    if s.len() == 8 {
        // "FFRRGGBB" -> "#RRGGBB"
        format!("#{}", &s[2..])
    } else if s.len() == 6 {
        format!("#{s}")
    } else if s.starts_with('#') {
        s.to_string()
    } else {
        format!("#{s}")
    }
}

/// Convert "#RRGGBB" to "FFRRGGBB" ARGB.
pub(super) fn hex_to_argb(hex: &str) -> String {
    let s = hex.strip_prefix('#').unwrap_or(hex);
    format!("FF{s}")
}

/// Map umya border style string to our canonical style names.
pub(super) fn umya_border_style_to_str(style: &str) -> &'static str {
    match style.to_ascii_lowercase().as_str() {
        "thin" => "thin",
        "medium" => "medium",
        "thick" => "thick",
        "double" => "double",
        "dashed" => "dashed",
        "dotted" => "dotted",
        "hair" => "hair",
        "mediumdashed" => "mediumDashed",
        "dashdot" => "dashDot",
        "mediumdashdot" => "mediumDashDot",
        "dashdotdot" => "dashDotDot",
        "mediumdashdotdot" => "mediumDashDotDot",
        "slantdashdot" => "slantDashDot",
        _ => "none",
    }
}

pub(super) fn col_letter_to_u32(col_str: &str) -> Result<u32, String> {
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(format!("Invalid column string: {col_str}"));
        }
        let uc = ch.to_ascii_uppercase() as u8;
        col = col * 26 + (uc - b'A' + 1) as u32;
    }
    if col == 0 {
        return Err(format!("Invalid column string: {col_str}"));
    }
    Ok(col)
}
