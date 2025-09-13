pub fn should_show(toggled_on: bool, high_contrast: bool, fullscreen: bool) -> bool {
    toggled_on && !high_contrast && !fullscreen
}

pub fn calc_top_center(
    work: (i32, i32, i32, i32),
    text_w: i32,
    _text_h: i32,
    margin: i32,
) -> (i32, i32) {
    let (left, top, right, _bottom) = work;
    let work_w = right - left;
    let x = left + (work_w - text_w) / 2;
    let y = top + margin;
    (x, y)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn center_calc_basic() {
        let (x, y) = calc_top_center((0, 0, 1000, 800), 200, 20, 10);
        assert_eq!(x, 400);
        assert_eq!(y, 10);
    }

    #[test]
    fn center_calc_nonzero_origin() {
        let (x, y) = calc_top_center((100, 50, 900, 1050), 300, 20, 8);
        // work width = 800; (800-300)/2 = 250; x = 100+250 = 350
        assert_eq!(x, 350);
        assert_eq!(y, 58);
    }
}
