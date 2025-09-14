fn make_icon() -> std::path::PathBuf {
    use ico::{IconDir, IconDirEntry, IconImage};
    use image::{ImageBuffer, Rgba};
    use std::env;
    use std::fs;
    use std::path::PathBuf;

    let out_dir = PathBuf::from(env::var("OUT_DIR").unwrap());
    let ico_path = out_dir.join("app.ico");

    // Generate a simple RGBA icon at multiple sizes.
    let mut dir = IconDir::new(ico::ResourceType::Icon);
    for &size in &[16u32, 32, 48, 64, 128, 256] {
        let mut img: ImageBuffer<Rgba<u8>, Vec<u8>> = ImageBuffer::new(size, size);
        // Background: dark blue with rounded-ish corners using simple distance field
        let r = (size as f32) * 0.2;
        for y in 0..size {
            for x in 0..size {
                let xf = x as f32 + 0.5;
                let yf = y as f32 + 0.5;
                let w = size as f32;
                let h = size as f32;
                // Distance to a rounded-rect border
                let dx = (xf - r).max(0.0) + (xf - (w - r)).min(0.0).abs();
                let dy = (yf - r).max(0.0) + (yf - (h - r)).min(0.0).abs();
                let inside = dx <= 0.0 || dy <= 0.0 || (dx * dx + dy * dy) <= r * r;
                if inside {
                    img.put_pixel(x, y, Rgba([18, 32, 64, 255]));
                } else {
                    img.put_pixel(x, y, Rgba([0, 0, 0, 0]));
                }
            }
        }
        // Simple horizontal highlight stripe
        for y in (size / 4)..(size / 4 + size / 8) {
            for x in 0..size {
                let px = img.get_pixel_mut(x, y);
                // Blend a lighter blue (avoid unnecessary min per Clippy)
                let a = 200u8;
                *px = Rgba([
                    (px[0] as u16 + 40) as u8,
                    (px[1] as u16 + 40) as u8,
                    (px[2] as u16 + 40) as u8,
                    a,
                ]);
            }
        }
        let rgba = img.into_raw();
        let icon_image = IconImage::from_rgba_data(size, size, rgba);
        let entry = IconDirEntry::encode(&icon_image).expect("encode icon entry");
        dir.add_entry(entry);
    }
    let mut file = fs::File::create(&ico_path).expect("create ico");
    dir.write(&mut file).expect("write ico");
    ico_path
}

fn main() {
    // Embed a manifest enabling Per-Monitor v2 DPI awareness.
    #[allow(unused_must_use)]
    {
        embed_manifest::embed_manifest_file("app.manifest");
    }

    // Generate an icon at build time and compile Windows version resources (icon + version info).
    #[cfg(windows)]
    {
        let ico_path = make_icon();
        let mut res = winres::WindowsResource::new();
        res.set_icon(&ico_path.to_string_lossy());

        // Version/info resources
        let pkg_ver = std::env::var("CARGO_PKG_VERSION").unwrap_or_else(|_| "1.0.0".into());
        let file_ver = format!("{}.0", pkg_ver); // Windows expects 4-part versions
        res.set("FileDescription", "Desktop Labeler");
        res.set("ProductName", "Desktop Labeler");
        res.set("CompanyName", "0x4D44 Software");
        res.set("FileVersion", &file_ver);
        res.set("ProductVersion", &file_ver);
        res.set("InternalName", "mddsklbl");
        res.set("OriginalFilename", "mddsklbl.exe");
        res.set("Comments", "Repo: mddsklbl");
        res.set("LegalCopyright", "(C) 2025 0x4D44 Software");

        res.compile().expect("failed to compile Windows resources");
    }
}
