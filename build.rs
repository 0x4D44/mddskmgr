fn main() {
    // Embed a manifest enabling Per-Monitor v2 DPI awareness.
    #[allow(unused_must_use)]
    {
        embed_manifest::embed_manifest_file("app.manifest");
    }
}
