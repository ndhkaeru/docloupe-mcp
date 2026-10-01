use std::env;
use std::fs;
use std::path::PathBuf;
use std::process::ExitCode;

use layout::backends::svg::SVGWriter;
use layout::gv::{self, parser::DotParser};
use layout::topo::layout::VisualGraph;

fn run() -> Result<(), String> {
    let mut arguments = env::args().skip(1);
    let operation = arguments.next().ok_or("expected validate or render")?;
    if operation != "validate" && operation != "render" {
        return Err("expected validate or render".into());
    }
    let mut language = None;
    let mut input = None;
    let mut output = None;
    while let Some(flag) = arguments.next() {
        let value = arguments
            .next()
            .ok_or(format!("missing value for {flag}"))?;
        match flag.as_str() {
            "--language" => language = Some(value),
            "--input" => input = Some(PathBuf::from(value)),
            "--output" => output = Some(PathBuf::from(value)),
            _ => return Err(format!("unknown option: {flag}")),
        }
    }
    let language = language.ok_or("missing --language")?;
    let input = input.ok_or("missing --input")?;
    let source = fs::read_to_string(input).map_err(|error| error.to_string())?;
    let svg = match language.as_str() {
        "mermaid" => {
            if operation == "validate" {
                mermaid_rs_renderer::parse_mermaid_strict(&source)
                    .map_err(|error| error.to_string())?;
                None
            } else {
                Some(
                    mermaid_rs_renderer::render_strict(
                        &source,
                        mermaid_rs_renderer::RenderOptions::default(),
                    )
                    .map_err(|error| error.to_string())?,
                )
            }
        }
        "dot" => {
            let mut parser = DotParser::new(&source);
            let graph = parser.process()?;
            if operation == "validate" {
                None
            } else {
                let mut builder = gv::GraphBuilder::new();
                builder.visit_graph(&graph);
                let mut visual: VisualGraph = builder.get();
                let mut writer = SVGWriter::new();
                visual.do_it(false, false, false, &mut writer);
                Some(writer.finalize())
            }
        }
        _ => return Err(format!("unsupported diagram language: {language}")),
    };
    if let Some(svg) = svg {
        let output = output.ok_or("render requires --output")?;
        fs::write(output, svg).map_err(|error| error.to_string())?;
    } else if output.is_some() {
        return Err("validate does not accept --output".into());
    }
    Ok(())
}

fn main() -> ExitCode {
    match run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            ExitCode::FAILURE
        }
    }
}
