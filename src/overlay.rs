use anyhow::{Result, anyhow};
use once_cell::sync::OnceCell;
use std::ffi::c_void;
use std::mem::{size_of, zeroed};
use windows::Win32::Foundation::{COLORREF, HWND, POINT, RECT, SIZE};
use windows::Win32::Graphics::Direct2D::Common::*;
use windows::Win32::Graphics::Direct2D::*;
use windows::Win32::Graphics::DirectWrite::*;
use windows::Win32::Graphics::Dxgi::Common::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::UI::WindowsAndMessaging::{
    HWND_TOPMOST, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SetWindowPos,
};
use windows::Win32::UI::WindowsAndMessaging::{
    SPI_GETWORKAREA, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS, SystemParametersInfoW, ULW_ALPHA,
    UpdateLayeredWindow,
};
use windows::core::{Interface, PCWSTR};

#[derive(Clone)]
pub struct Overlay {
    hwnd: HWND,
    font_family: String,
    font_px: i32,
}

impl Overlay {
    pub fn new(hwnd: HWND, font_family: &str, font_size_dip: u32) -> Result<Self> {
        // Approximate 1 DIP = 1 px at 100% scale for initial implementation
        Ok(Self {
            hwnd,
            font_family: font_family.to_string(),
            font_px: font_size_dip as i32,
        })
    }

    pub fn draw_line_top_center(&self, text: &str, margin_px: i32) -> Result<()> {
        self.draw_line_top_center_with_hints(text, "", margin_px)
    }

    pub fn draw_line_top_center_with_hints(
        &self,
        text: &str,
        hints: &str,
        margin_px: i32,
    ) -> Result<()> {
        tracing::debug!(text=%text, hints=%hints, "overlay: draw_line_top_center");
        let (w, h) = self.measure_text_with_hints(text, hints)?;
        let w_pad = w + margin_px * 2;
        let h_pad = h + margin_px * 2;

        // Compute top-center position on the primary work area
        let mut work: RECT = RECT::default();
        unsafe {
            let _ = SystemParametersInfoW(
                SPI_GETWORKAREA,
                0,
                Some(&mut work as *mut _ as *mut c_void),
                SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
            );
        };
        let work_w = work.right - work.left;
        let x = work.left + (work_w - w_pad) / 2;
        let y = work.top + margin_px;

        let res = self.render_and_update(text, hints, x, y, w_pad, h_pad, margin_px);
        if let Err(e) = &res {
            tracing::warn!(error=?e, "overlay: render_and_update error");
        }
        res
    }

    pub fn draw_line_top_anchor_with_hints(
        &self,
        text: &str,
        hints: &str,
        margin_px: i32,
        anchor_ratio: f32,
    ) -> Result<()> {
        let (w, h) = self.measure_text_with_hints(text, hints)?;
        let w_pad = w + margin_px * 2;
        let h_pad = h + margin_px * 2;
        // Work area
        let mut work: RECT = RECT::default();
        unsafe {
            let _ = SystemParametersInfoW(
                SPI_GETWORKAREA,
                0,
                Some(&mut work as *mut _ as *mut c_void),
                SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
            );
        };
        let work_w = work.right - work.left;
        let anchor_x = work.left as f32 + (work_w as f32 * anchor_ratio);
        let mut x = anchor_x.round() as i32 - w_pad / 2;
        if x < work.left {
            x = work.left;
        }
        if x + w_pad > work.right {
            x = work.right - w_pad;
        }
        let y = work.top + margin_px;
        let res = self.render_and_update(text, hints, x, y, w_pad, h_pad, margin_px);
        if let Err(e) = &res {
            tracing::warn!(error=?e, "overlay: render_and_update error");
        }
        res
    }

    #[allow(clippy::too_many_arguments)]
    fn render_and_update(
        &self,
        text: &str,
        hints: &str,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        pad: i32,
    ) -> Result<()> {
        unsafe {
            let screen_dc = GetDC(None);
            if screen_dc.0.is_null() {
                return Err(anyhow!("GetDC failed"));
            }
            let mem_dc = CreateCompatibleDC(screen_dc);
            if mem_dc.0.is_null() {
                ReleaseDC(None, screen_dc);
                return Err(anyhow!("CreateCompatibleDC failed"));
            }

            // Create top-down 32bpp DIB
            let mut bi: BITMAPINFO = zeroed();
            bi.bmiHeader.biSize = size_of::<BITMAPINFOHEADER>() as u32;
            bi.bmiHeader.biWidth = width;
            bi.bmiHeader.biHeight = -height; // top-down
            bi.bmiHeader.biPlanes = 1;
            bi.bmiHeader.biBitCount = 32;
            bi.bmiHeader.biCompression = BI_RGB.0;

            let mut bits: *mut c_void = std::ptr::null_mut();
            let hbm = CreateDIBSection(mem_dc, &bi, DIB_RGB_COLORS, &mut bits, None, 0)?;
            let old = SelectObject(mem_dc, HGDIOBJ(hbm.0));

            // Fill background (black) — no per-pixel alpha; use global alpha in blend
            let stride = (width * 4) as usize;
            let total = (height as usize) * stride;
            let buf = std::slice::from_raw_parts_mut(bits as *mut u8, total);
            for y in 0..height as usize {
                let row = &mut buf[y * stride..(y + 1) * stride];
                for px in row.chunks_exact_mut(4) {
                    // BGRA order; opaque RGB, alpha ignored because we use global SourceConstantAlpha
                    px[0] = 0; // B
                    px[1] = 0; // G
                    px[2] = 0; // R
                    px[3] = 0; // A (transparent; D2D will draw alpha)
                }
            }

            // Prefer Direct2D per-pixel alpha; fallback to GDI if it fails
            let d2d_ok = render_d2d_with_hints(
                mem_dc,
                width,
                height,
                pad,
                text,
                hints,
                &self.font_family,
                self.font_px,
            )
            .is_ok();
            if !d2d_ok {
                let font = create_font(&self.font_family, self.font_px);
                let old_font = SelectObject(mem_dc, HGDIOBJ(font.0));
                SetBkMode(mem_dc, TRANSPARENT);
                let white = COLORREF(0x00FFFFFF);
                let _ = SetTextColor(mem_dc, white);
                let mut rc = RECT {
                    left: pad,
                    top: pad,
                    right: width - pad,
                    bottom: height - pad,
                };
                let mut wtext: Vec<u16> = if hints.is_empty() {
                    text.encode_utf16().collect()
                } else {
                    format!("{} {}", text, hints).encode_utf16().collect()
                };
                let _ = DrawTextW(
                    mem_dc,
                    &mut wtext,
                    &mut rc,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX,
                );
                let _ = DeleteObject(HGDIOBJ(font.0));
                SelectObject(mem_dc, old_font);
            }

            // Apply a rounded window region to clip hit-testing and visuals
            let radius = (self.font_px / 2).clamp(6, 20);
            let hrgn = CreateRoundRectRgn(0, 0, width, height, radius * 2, radius * 2);
            let _ = SetWindowRgn(self.hwnd, hrgn, true);

            // Push to layered window with uniform alpha
            let src_pt = POINT { x: 0, y: 0 };
            let dst_pt = POINT { x, y };
            let size = SIZE {
                cx: width,
                cy: height,
            };
            let alpha_format = if d2d_ok { 1u8 } else { 0u8 };
            let src_const = if d2d_ok { 255u8 } else { 200u8 };
            let blend = BLENDFUNCTION {
                BlendOp: 0u8,
                BlendFlags: 0u8,
                SourceConstantAlpha: src_const,
                AlphaFormat: alpha_format,
            };
            let ulw_res = UpdateLayeredWindow(
                self.hwnd,
                HDC(std::ptr::null_mut()),
                Some(&dst_pt),
                Some(&size),
                mem_dc,
                Some(&src_pt),
                COLORREF(0),
                Some(&blend),
                ULW_ALPHA,
            );
            if let Err(e) = &ulw_res {
                tracing::warn!(error=?e, "overlay: UpdateLayeredWindow failed");
            }

            // Reassert topmost after painting without activating
            let _ = SetWindowPos(
                self.hwnd,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );

            // Cleanup
            SelectObject(mem_dc, old);
            let _ = DeleteObject(HGDIOBJ(hbm.0));
            let _ = DeleteDC(mem_dc);
            ReleaseDC(None, screen_dc);
        }
        Ok(())
    }

    fn measure_text_with_hints(&self, text: &str, hints: &str) -> Result<(i32, i32)> {
        // Use DirectWrite for accurate measurement (apply smaller font to hints)
        let factory = get_dwrite_factory()?;
        unsafe {
            let tf = factory.CreateTextFormat(
                PCWSTR(to_utf16(&self.font_family).as_ptr()),
                None,
                DWRITE_FONT_WEIGHT_NORMAL,
                DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL,
                self.font_px as f32,
                PCWSTR(windows::core::w!("en-US").as_wide().as_ptr()),
            )?;
            let combined = if hints.is_empty() {
                text.to_string()
            } else {
                format!("{} {}", text, hints)
            };
            let s = to_utf16(&combined);
            let layout = factory.CreateTextLayout(&s[..s.len() - 1], &tf, 4096.0, 4096.0)?;
            if !hints.is_empty() {
                let hint_start = text.encode_utf16().count() + 1; // +1 for the space
                let hint_len = hints.encode_utf16().count();
                let small = (self.font_px as f32 * 0.7).max(8.0);
                let range = DWRITE_TEXT_RANGE {
                    startPosition: hint_start as u32,
                    length: hint_len as u32,
                };
                let _ = layout.SetFontSize(small, range);
            }
            let mut m = DWRITE_TEXT_METRICS::default();
            layout.GetMetrics(&mut m)?;
            let w = m.widthIncludingTrailingWhitespace.ceil() as i32;
            let h = m.height.ceil() as i32;
            Ok((w.max(1), h.max(1)))
        }
    }
}

fn create_font(face: &str, px: i32) -> HFONT {
    let height = -px; // negative height means character height in logical units
    let wface = to_utf16(face);
    unsafe {
        CreateFontW(
            height,
            0,
            0,
            0,
            FW_NORMAL.0 as i32,
            0,
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            0,
            0,
            CLEARTYPE_QUALITY.0 as u32,
            DEFAULT_PITCH.0 as u32,
            PCWSTR(wface.as_ptr()),
        )
    }
}

fn to_utf16(s: &str) -> Vec<u16> {
    let mut v: Vec<u16> = s.encode_utf16().collect();
    v.push(0);
    v
}

fn get_dwrite_factory() -> Result<&'static IDWriteFactory> {
    static FACTORY: OnceCell<IDWriteFactory> = OnceCell::new();
    FACTORY.get_or_try_init(|| {
        unsafe { DWriteCreateFactory::<IDWriteFactory>(DWRITE_FACTORY_TYPE_SHARED) }
            .map_err(|e| anyhow!("DWriteCreateFactory failed: {:?}", e))
    })
}

fn get_d2d_factory() -> Result<&'static ID2D1Factory> {
    static FACTORY: OnceCell<ID2D1Factory> = OnceCell::new();
    FACTORY.get_or_try_init(|| {
        unsafe { D2D1CreateFactory::<ID2D1Factory>(D2D1_FACTORY_TYPE_SINGLE_THREADED, None) }
            .map_err(|e| anyhow!("D2D1CreateFactory failed: {:?}", e))
    })
}

#[allow(clippy::too_many_arguments)]
fn render_d2d_with_hints(
    hdc: HDC,
    width: i32,
    height: i32,
    pad: i32,
    text: &str,
    hints: &str,
    font: &str,
    font_px: i32,
) -> Result<()> {
    let factory = get_d2d_factory()?;
    unsafe {
        let props = D2D1_RENDER_TARGET_PROPERTIES {
            r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
            pixelFormat: D2D1_PIXEL_FORMAT {
                format: DXGI_FORMAT_B8G8R8A8_UNORM,
                alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
            },
            dpiX: 0.0,
            dpiY: 0.0,
            usage: D2D1_RENDER_TARGET_USAGE_NONE,
            minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
        };
        let rt = factory.CreateDCRenderTarget(&props)?;
        let rc = RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: height,
        };
        rt.BindDC(hdc, &rc)?;
        rt.BeginDraw();

        rt.Clear(Some(&D2D1_COLOR_F {
            r: 0.0,
            g: 0.0,
            b: 0.0,
            a: 0.0,
        }));

        let base: ID2D1RenderTarget = rt.cast()?;
        let bg = base.CreateSolidColorBrush(
            &D2D1_COLOR_F {
                r: 0.0,
                g: 0.0,
                b: 0.0,
                a: 0.35,
            },
            None,
        )?;
        let rounded = D2D1_ROUNDED_RECT {
            rect: D2D_RECT_F {
                left: 0.0,
                top: 0.0,
                right: width as f32,
                bottom: height as f32,
            },
            radiusX: (font_px as f32 * 0.5).clamp(6.0, 20.0),
            radiusY: (font_px as f32 * 0.5).clamp(6.0, 20.0),
        };
        base.FillRoundedRectangle(&rounded, &bg);

        let dwrite = get_dwrite_factory()?;
        let tf = dwrite.CreateTextFormat(
            PCWSTR(to_utf16(font).as_ptr()),
            None,
            DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            font_px as f32,
            PCWSTR(windows::core::w!("en-US").as_wide().as_ptr()),
        )?;
        let combined = if hints.is_empty() {
            text.to_string()
        } else {
            format!("{} {}", text, hints)
        };
        let s16 = to_utf16(&combined);
        let layout = dwrite.CreateTextLayout(
            &s16[..s16.len() - 1],
            &tf,
            (width - pad) as f32,
            (height - pad) as f32,
        )?;
        if !hints.is_empty() {
            let hint_start = text.encode_utf16().count() + 1;
            let hint_len = hints.encode_utf16().count();
            let small = (font_px as f32 * 0.7).max(8.0);
            let range = DWRITE_TEXT_RANGE {
                startPosition: hint_start as u32,
                length: hint_len as u32,
            };
            let _ = layout.SetFontSize(small, range);
        }
        let fg = base.CreateSolidColorBrush(
            &D2D1_COLOR_F {
                r: 1.0,
                g: 1.0,
                b: 1.0,
                a: 1.0,
            },
            None,
        )?;
        let origin = D2D_POINT_2F {
            x: pad as f32,
            y: pad as f32,
        };
        base.DrawTextLayout(origin, &layout, &fg, D2D1_DRAW_TEXT_OPTIONS_NONE);

        base.EndDraw(None, None)?;
    }
    Ok(())
}
