use md2docx::mermaid::{render_to_png, render_to_svg};

fn main() {
    let diagrams = vec![
        ("flowchart", "flowchart LR\n    A[Start] --> B[End]"),
        ("sequence", "sequenceDiagram\n    A->>B: Hello"),
        ("state", "stateDiagram-v2\n    [*] --> Idle"),
    ];

    for (name, content) in diagrams {
        match render_to_svg(content) {
            Ok(svg) => {
                println!("=== {} SVG ===", name);
                println!("Length: {} bytes", svg.len());

                // Try PNG
                match render_to_png(content, 2.0) {
                    Ok(png) => {
                        println!("{} PNG Success: {} bytes", name, png.len());
                        std::fs::write(format!("/tmp/{}_debug.png", name), png).unwrap();
                    }
                    Err(e) => println!("{} PNG Failed: {}", name, e),
                }
            }
            Err(e) => println!("{} SVG failed: {}", name, e),
        }
    }
}
