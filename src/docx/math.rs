//! LaTeX math to Office Math Markup Language (OMML) converter
//!
//! Converts LaTeX math expressions into OMML XML for embedding in DOCX documents.
//! Supports common LaTeX constructs: fractions, superscripts, subscripts, Greek letters,
//! operators, roots, summations, integrals, and more.

use std::collections::HashMap;
use once_cell::sync::Lazy;

/// Map of LaTeX command names to Unicode characters
static LATEX_SYMBOLS: Lazy<HashMap<&'static str, &'static str>> = Lazy::new(|| {
    let mut m = HashMap::new();
    // Greek lowercase
    m.insert("alpha", "\u{03B1}");
    m.insert("beta", "\u{03B2}");
    m.insert("gamma", "\u{03B3}");
    m.insert("delta", "\u{03B4}");
    m.insert("epsilon", "\u{03B5}");
    m.insert("varepsilon", "\u{03B5}");
    m.insert("zeta", "\u{03B6}");
    m.insert("eta", "\u{03B7}");
    m.insert("theta", "\u{03B8}");
    m.insert("vartheta", "\u{03D1}");
    m.insert("iota", "\u{03B9}");
    m.insert("kappa", "\u{03BA}");
    m.insert("lambda", "\u{03BB}");
    m.insert("mu", "\u{03BC}");
    m.insert("nu", "\u{03BD}");
    m.insert("xi", "\u{03BE}");
    m.insert("pi", "\u{03C0}");
    m.insert("rho", "\u{03C1}");
    m.insert("sigma", "\u{03C3}");
    m.insert("tau", "\u{03C4}");
    m.insert("upsilon", "\u{03C5}");
    m.insert("phi", "\u{03C6}");
    m.insert("varphi", "\u{03D5}");
    m.insert("chi", "\u{03C7}");
    m.insert("psi", "\u{03C8}");
    m.insert("omega", "\u{03C9}");
    // Greek uppercase
    m.insert("Gamma", "\u{0393}");
    m.insert("Delta", "\u{0394}");
    m.insert("Theta", "\u{0398}");
    m.insert("Lambda", "\u{039B}");
    m.insert("Xi", "\u{039E}");
    m.insert("Pi", "\u{03A0}");
    m.insert("Sigma", "\u{03A3}");
    m.insert("Upsilon", "\u{03A5}");
    m.insert("Phi", "\u{03A6}");
    m.insert("Psi", "\u{03A8}");
    m.insert("Omega", "\u{03A9}");
    // Operators and relations
    m.insert("times", "\u{00D7}");
    m.insert("div", "\u{00F7}");
    m.insert("cdot", "\u{22C5}");
    m.insert("pm", "\u{00B1}");
    m.insert("mp", "\u{2213}");
    m.insert("leq", "\u{2264}");
    m.insert("le", "\u{2264}");
    m.insert("geq", "\u{2265}");
    m.insert("ge", "\u{2265}");
    m.insert("neq", "\u{2260}");
    m.insert("ne", "\u{2260}");
    m.insert("approx", "\u{2248}");
    m.insert("equiv", "\u{2261}");
    m.insert("sim", "\u{223C}");
    m.insert("propto", "\u{221D}");
    m.insert("infty", "\u{221E}");
    m.insert("partial", "\u{2202}");
    m.insert("nabla", "\u{2207}");
    m.insert("forall", "\u{2200}");
    m.insert("exists", "\u{2203}");
    m.insert("in", "\u{2208}");
    m.insert("notin", "\u{2209}");
    m.insert("subset", "\u{2282}");
    m.insert("supset", "\u{2283}");
    m.insert("subseteq", "\u{2286}");
    m.insert("supseteq", "\u{2287}");
    m.insert("cup", "\u{222A}");
    m.insert("cap", "\u{2229}");
    m.insert("emptyset", "\u{2205}");
    m.insert("varnothing", "\u{2205}");
    // Arrows
    m.insert("rightarrow", "\u{2192}");
    m.insert("to", "\u{2192}");
    m.insert("leftarrow", "\u{2190}");
    m.insert("leftrightarrow", "\u{2194}");
    m.insert("Rightarrow", "\u{21D2}");
    m.insert("Leftarrow", "\u{21D0}");
    m.insert("Leftrightarrow", "\u{21D4}");
    m.insert("implies", "\u{21D2}");
    m.insert("iff", "\u{21D4}");
    // Misc
    m.insert("ldots", "\u{2026}");
    m.insert("cdots", "\u{22EF}");
    m.insert("vdots", "\u{22EE}");
    m.insert("ddots", "\u{22F1}");
    m.insert("therefore", "\u{2234}");
    m.insert("because", "\u{2235}");
    m.insert("angle", "\u{2220}");
    m.insert("perp", "\u{22A5}");
    m.insert("parallel", "\u{2225}");
    m.insert("star", "\u{22C6}");
    m.insert("circ", "\u{2218}");
    m.insert("bullet", "\u{2022}");
    m.insert("neg", "\u{00AC}");
    m.insert("lnot", "\u{00AC}");
    m.insert("land", "\u{2227}");
    m.insert("lor", "\u{2228}");
    m.insert("wedge", "\u{2227}");
    m.insert("vee", "\u{2228}");
    m.insert("oplus", "\u{2295}");
    m.insert("otimes", "\u{2297}");
    // Spacing
    m.insert("quad", "\u{2003}");
    m.insert("qquad", "\u{2003}\u{2003}");
    m.insert(",", "\u{2009}");
    m.insert(";", "\u{2005}");
    m.insert("!", "");
    m.insert(" ", " ");
    // Text-mode accents (used in math)
    m.insert("hat", "\u{0302}"); // placeholder, handled specially
    m.insert("bar", "\u{0304}");
    m.insert("dot", "\u{0307}");
    m.insert("ddot", "\u{0308}");
    m.insert("tilde", "\u{0303}");
    m.insert("vec", "\u{20D7}");
    // Functions (handled as operator names)
    m.insert("sin", "sin");
    m.insert("cos", "cos");
    m.insert("tan", "tan");
    m.insert("cot", "cot");
    m.insert("sec", "sec");
    m.insert("csc", "csc");
    m.insert("arcsin", "arcsin");
    m.insert("arccos", "arccos");
    m.insert("arctan", "arctan");
    m.insert("sinh", "sinh");
    m.insert("cosh", "cosh");
    m.insert("tanh", "tanh");
    m.insert("log", "log");
    m.insert("ln", "ln");
    m.insert("exp", "exp");
    m.insert("lim", "lim");
    m.insert("limsup", "lim sup");
    m.insert("liminf", "lim inf");
    m.insert("sup", "sup");
    m.insert("inf", "inf");
    m.insert("min", "min");
    m.insert("max", "max");
    m.insert("det", "det");
    m.insert("dim", "dim");
    m.insert("ker", "ker");
    m.insert("hom", "hom");
    m.insert("deg", "deg");
    m.insert("gcd", "gcd");
    m.insert("arg", "arg");
    m.insert("mod", "mod");
    m.insert("Pr", "Pr");
    m
});

/// Set of LaTeX commands that are function names (rendered upright, not italic)
static FUNCTION_NAMES: Lazy<std::collections::HashSet<&'static str>> = Lazy::new(|| {
    [
        "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
        "sinh", "cosh", "tanh", "log", "ln", "exp", "lim", "limsup", "liminf",
        "sup", "inf", "min", "max", "det", "dim", "ker", "hom", "deg", "gcd",
        "arg", "mod", "Pr",
    ]
    .into_iter()
    .collect()
});

/// Nary operators (large operators with limits)
static NARY_OPERATORS: Lazy<HashMap<&'static str, &'static str>> = Lazy::new(|| {
    let mut m = HashMap::new();
    m.insert("sum", "\u{2211}");
    m.insert("prod", "\u{220F}");
    m.insert("coprod", "\u{2210}");
    m.insert("int", "\u{222B}");
    m.insert("iint", "\u{222C}");
    m.insert("iiint", "\u{222D}");
    m.insert("oint", "\u{222E}");
    m.insert("bigcup", "\u{22C3}");
    m.insert("bigcap", "\u{22C2}");
    m.insert("bigoplus", "\u{2A01}");
    m.insert("bigotimes", "\u{2A02}");
    m
});

/// Convert a LaTeX math expression to OMML XML string.
///
/// This produces the inner content of an `<m:oMath>` element.
pub fn latex_to_omml(latex: &str) -> String {
    let tokens = tokenize(latex);
    let mut output = String::new();
    tokens_to_omml(&tokens, &mut output);
    output
}

/// Convert a LaTeX math expression to a complete `<m:oMathPara>` block for display math.
pub fn latex_to_omml_paragraph(latex: &str) -> String {
    let inner = latex_to_omml(latex);
    format!(
        "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr><m:oMath>{}</m:oMath></m:oMathPara>",
        inner
    )
}

/// Convert a LaTeX math expression to an inline `<m:oMath>` element.
pub fn latex_to_omml_inline(latex: &str) -> String {
    let inner = latex_to_omml(latex);
    format!("<m:oMath>{}</m:oMath>", inner)
}

// ── Tokenizer ──────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Text(String),         // plain characters
    Command(String),      // \command
    Group(Vec<Token>),    // { ... }
    Superscript,          // ^
    Subscript,            // _
    Ampersand,            // & (for matrices)
    Newline,              // \\ (for matrices)
    OpenBracket,          // [
    CloseBracket,         // ]
}

fn tokenize(input: &str) -> Vec<Token> {
    let chars: Vec<char> = input.chars().collect();
    let mut pos = 0;
    tokenize_until(&chars, &mut pos, None)
}

fn tokenize_until(chars: &[char], pos: &mut usize, end_char: Option<char>) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut text_buf = String::new();

    while *pos < chars.len() {
        let ch = chars[*pos];

        if Some(ch) == end_char {
            if !text_buf.is_empty() {
                tokens.push(Token::Text(text_buf.clone()));
                text_buf.clear();
            }
            *pos += 1;
            return tokens;
        }

        match ch {
            '\\' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                *pos += 1;
                if *pos < chars.len() {
                    let next = chars[*pos];
                    if next == '\\' {
                        tokens.push(Token::Newline);
                        *pos += 1;
                    } else if next == '{' || next == '}' || next == '|' || next == '&'
                        || next == '%' || next == '#' || next == '_'
                    {
                        tokens.push(Token::Text(next.to_string()));
                        *pos += 1;
                    } else if next == ' ' || next == ',' || next == ';' || next == '!' {
                        // Spacing commands
                        let cmd = next.to_string();
                        if let Some(sym) = LATEX_SYMBOLS.get(cmd.as_str()) {
                            tokens.push(Token::Text(sym.to_string()));
                        }
                        *pos += 1;
                    } else if next.is_alphabetic() {
                        let mut cmd = String::new();
                        while *pos < chars.len() && chars[*pos].is_alphabetic() {
                            cmd.push(chars[*pos]);
                            *pos += 1;
                        }
                        tokens.push(Token::Command(cmd));
                    } else {
                        tokens.push(Token::Text(next.to_string()));
                        *pos += 1;
                    }
                }
            }
            '{' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                *pos += 1;
                let group = tokenize_until(chars, pos, Some('}'));
                tokens.push(Token::Group(group));
            }
            '^' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                tokens.push(Token::Superscript);
                *pos += 1;
            }
            '_' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                tokens.push(Token::Subscript);
                *pos += 1;
            }
            '&' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                tokens.push(Token::Ampersand);
                *pos += 1;
            }
            '[' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                tokens.push(Token::OpenBracket);
                *pos += 1;
            }
            ']' => {
                if !text_buf.is_empty() {
                    tokens.push(Token::Text(text_buf.clone()));
                    text_buf.clear();
                }
                tokens.push(Token::CloseBracket);
                *pos += 1;
            }
            ' ' | '\t' | '\n' | '\r' => {
                // Skip whitespace in math mode (spacing is semantic via commands)
                *pos += 1;
            }
            _ => {
                text_buf.push(ch);
                *pos += 1;
            }
        }
    }

    if !text_buf.is_empty() {
        tokens.push(Token::Text(text_buf));
    }

    tokens
}

// ── OMML Generator ─────────────────────────────────────────────────────

fn tokens_to_omml(tokens: &[Token], output: &mut String) {
    let mut i = 0;
    while i < tokens.len() {
        match &tokens[i] {
            Token::Command(cmd) => {
                i += 1;
                i += handle_command(cmd, &tokens[i..], output);
            }
            Token::Text(text) => {
                // Check if next tokens create a sub/superscript
                if i + 2 < tokens.len() && (tokens[i + 1] == Token::Subscript || tokens[i + 1] == Token::Superscript) {
                    // Let the sub/superscript handler deal with collecting the base
                    emit_text_with_scripts(text, &tokens[i + 1..], output, &mut i);
                } else {
                    write_run(output, text, false);
                    i += 1;
                }
            }
            Token::Group(inner) => {
                // Check if next token is sub/superscript
                if i + 1 < tokens.len() && (tokens[i + 1] == Token::Subscript || tokens[i + 1] == Token::Superscript) {
                    let mut base_xml = String::new();
                    tokens_to_omml(inner, &mut base_xml);
                    emit_raw_with_scripts(&base_xml, &tokens[i + 1..], output, &mut i);
                } else {
                    tokens_to_omml(inner, output);
                    i += 1;
                }
            }
            Token::Superscript => {
                // Bare superscript without base (shouldn't happen normally)
                if i + 1 < tokens.len() {
                    let sup = &tokens[i + 1];
                    let mut sup_xml = String::new();
                    token_arg_to_omml(sup, &mut sup_xml);
                    output.push_str(&format!(
                        "<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>{}</m:e><m:sup>{}</m:sup></m:sSup>",
                        "", sup_xml
                    ));
                    i += 2;
                } else {
                    i += 1;
                }
            }
            Token::Subscript => {
                // Bare subscript without base
                if i + 1 < tokens.len() {
                    let sub = &tokens[i + 1];
                    let mut sub_xml = String::new();
                    token_arg_to_omml(sub, &mut sub_xml);
                    output.push_str(&format!(
                        "<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>{}</m:e><m:sub>{}</m:sub></m:sSub>",
                        "", sub_xml
                    ));
                    i += 2;
                } else {
                    i += 1;
                }
            }
            Token::Ampersand | Token::Newline | Token::OpenBracket | Token::CloseBracket => {
                i += 1;
            }
        }
    }
}

/// Handle a LaTeX command and return how many additional tokens were consumed
fn handle_command(cmd: &str, rest: &[Token], output: &mut String) -> usize {
    let mut consumed = 0;

    // Check for nary operators (sum, int, prod, etc.)
    if let Some(chr) = NARY_OPERATORS.get(cmd) {
        // Look for subscript/superscript limits
        let mut sub_xml = String::new();
        let mut sup_xml = String::new();
        let mut rest_idx = 0;

        // Consume sub/super in any order
        for _ in 0..2 {
            if rest_idx < rest.len() {
                if rest[rest_idx] == Token::Subscript {
                    rest_idx += 1;
                    if rest_idx < rest.len() {
                        let arg = &rest[rest_idx];
                        token_arg_to_omml(arg, &mut sub_xml);
                        rest_idx += 1;
                    }
                } else if rest[rest_idx] == Token::Superscript {
                    rest_idx += 1;
                    if rest_idx < rest.len() {
                        let arg = &rest[rest_idx];
                        token_arg_to_omml(arg, &mut sup_xml);
                        rest_idx += 1;
                    }
                }
            }
        }

        output.push_str("<m:nary>");
        output.push_str("<m:naryPr>");
        output.push_str(&format!("<m:chr m:val=\"{}\"/>", xml_escape(chr)));
        if sub_xml.is_empty() && sup_xml.is_empty() {
            output.push_str("<m:limLoc m:val=\"undOvr\"/>");
            output.push_str("<m:subHide m:val=\"1\"/>");
            output.push_str("<m:supHide m:val=\"1\"/>");
        } else {
            output.push_str("<m:limLoc m:val=\"subSup\"/>");
            if sub_xml.is_empty() {
                output.push_str("<m:subHide m:val=\"1\"/>");
            }
            if sup_xml.is_empty() {
                output.push_str("<m:supHide m:val=\"1\"/>");
            }
        }
        output.push_str("<m:ctrlPr/></m:naryPr>");
        output.push_str(&format!("<m:sub>{}</m:sub>", sub_xml));
        output.push_str(&format!("<m:sup>{}</m:sup>", sup_xml));

        // The "body" of the nary: collect remaining tokens as the body
        // In practice, the body extends to the end of the group or expression
        // For simplicity, take the next token/group as the body
        if rest_idx < rest.len() {
            let mut body_xml = String::new();
            let arg = &rest[rest_idx];
            token_arg_to_omml(arg, &mut body_xml);
            output.push_str(&format!("<m:e>{}</m:e>", body_xml));
            rest_idx += 1;
        } else {
            output.push_str("<m:e/>");
        }

        output.push_str("</m:nary>");
        return rest_idx;
    }

    // Fractions: \frac{num}{den}
    if cmd == "frac" || cmd == "dfrac" || cmd == "tfrac" {
        if rest.len() >= 2 {
            let mut num_xml = String::new();
            let mut den_xml = String::new();
            token_arg_to_omml(&rest[0], &mut num_xml);
            token_arg_to_omml(&rest[1], &mut den_xml);
            output.push_str(&format!(
                "<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>{}</m:num><m:den>{}</m:den></m:f>",
                num_xml, den_xml
            ));
            return 2;
        }
    }

    // Square root: \sqrt{x} or \sqrt[n]{x}
    if cmd == "sqrt" {
        if !rest.is_empty() {
            // Check for optional argument [n]
            if rest[0] == Token::OpenBracket {
                // Find the closing bracket to get the index
                let mut bracket_content = Vec::new();
                let mut j = 1;
                while j < rest.len() && rest[j] != Token::CloseBracket {
                    bracket_content.push(rest[j].clone());
                    j += 1;
                }
                if j < rest.len() {
                    j += 1; // skip ']'
                }
                if j < rest.len() {
                    let mut deg_xml = String::new();
                    let mut body_xml = String::new();
                    tokens_to_omml(&bracket_content, &mut deg_xml);
                    token_arg_to_omml(&rest[j], &mut body_xml);
                    output.push_str(&format!(
                        "<m:rad><m:radPr><m:ctrlPr/></m:radPr><m:deg>{}</m:deg><m:e>{}</m:e></m:rad>",
                        deg_xml, body_xml
                    ));
                    return j + 1;
                }
            }
            // Simple sqrt
            let mut body_xml = String::new();
            token_arg_to_omml(&rest[0], &mut body_xml);
            output.push_str(&format!(
                "<m:rad><m:radPr><m:degHide m:val=\"1\"/><m:ctrlPr/></m:radPr><m:deg/><m:e>{}</m:e></m:rad>",
                body_xml
            ));
            return 1;
        }
    }

    // Overline/bar: \overline{x} or \bar{x}
    if cmd == "overline" || cmd == "bar" || cmd == "hat" || cmd == "tilde"
        || cmd == "vec" || cmd == "dot" || cmd == "ddot" || cmd == "widehat" || cmd == "widetilde"
    {
        let accent_char = match cmd {
            "overline" | "bar" => "\u{0305}",
            "hat" | "widehat" => "\u{0302}",
            "tilde" | "widetilde" => "\u{0303}",
            "vec" => "\u{20D7}",
            "dot" => "\u{0307}",
            "ddot" => "\u{0308}",
            _ => "\u{0305}",
        };
        if !rest.is_empty() {
            let mut body_xml = String::new();
            token_arg_to_omml(&rest[0], &mut body_xml);
            output.push_str(&format!(
                "<m:acc><m:accPr><m:chr m:val=\"{}\"/><m:ctrlPr/></m:accPr><m:e>{}</m:e></m:acc>",
                accent_char, body_xml
            ));
            return 1;
        }
    }

    // Matrix/pmatrix/bmatrix environments
    if cmd == "begin" {
        if !rest.is_empty() {
            if let Token::Group(env_tokens) = &rest[0] {
                let env_name = tokens_to_text(env_tokens);
                consumed = 1;
                // Find matching \end{env_name}
                let (body_tokens, end_consumed) = find_env_body(&rest[consumed..], &env_name);
                consumed += end_consumed;

                match env_name.as_str() {
                    "matrix" | "pmatrix" | "bmatrix" | "Bmatrix" | "vmatrix" | "Vmatrix" => {
                        let (beg_chr, end_chr) = match env_name.as_str() {
                            "pmatrix" => ("(", ")"),
                            "bmatrix" => ("[", "]"),
                            "Bmatrix" => ("{", "}"),
                            "vmatrix" => ("|", "|"),
                            "Vmatrix" => ("\u{2016}", "\u{2016}"),
                            _ => ("", ""), // plain matrix
                        };
                        let rows = parse_matrix_body(&body_tokens);
                        emit_matrix(output, &rows, beg_chr, end_chr);
                    }
                    "cases" => {
                        let rows = parse_matrix_body(&body_tokens);
                        emit_matrix(output, &rows, "{", "");
                    }
                    "aligned" | "align" | "align*" => {
                        // Treat as equation array
                        let rows = parse_matrix_body(&body_tokens);
                        emit_equation_array(output, &rows);
                    }
                    _ => {
                        // Unknown environment, just process body
                        tokens_to_omml(&body_tokens, output);
                    }
                }
                return consumed;
            }
        }
    }

    // \left and \right delimiters
    if cmd == "left" {
        let (open_delim, body_tokens, close_delim, total) = parse_left_right(rest);
        let mut body_xml = String::new();
        tokens_to_omml(&body_tokens, &mut body_xml);
        output.push_str(&format!(
            "<m:d><m:dPr><m:begChr m:val=\"{}\"/><m:endChr m:val=\"{}\"/><m:ctrlPr/></m:dPr><m:e>{}</m:e></m:d>",
            xml_escape(&open_delim), xml_escape(&close_delim), body_xml
        ));
        return total;
    }

    // \text{...}
    if cmd == "text" || cmd == "textrm" || cmd == "textbf" || cmd == "textit" || cmd == "mathrm" || cmd == "mathbf" || cmd == "mathit" || cmd == "mathbb" || cmd == "mathcal" || cmd == "mathsf" {
        if !rest.is_empty() {
            let text_content = token_to_text(&rest[0]);
            let is_normal = matches!(cmd, "text" | "textrm" | "mathrm");
            if is_normal {
                write_run(output, &text_content, true);
            } else {
                write_run(output, &text_content, false);
            }
            return 1;
        }
    }

    // \operatorname{...}
    if cmd == "operatorname" {
        if !rest.is_empty() {
            let name = token_to_text(&rest[0]);
            write_run(output, &name, true);
            return 1;
        }
    }

    // Function names (sin, cos, log, etc.)
    if FUNCTION_NAMES.contains(cmd) {
        // Check if followed by sub/superscript (e.g. \lim_{x \to 0})
        let func_text = LATEX_SYMBOLS.get(cmd).copied().unwrap_or(cmd);
        if !rest.is_empty() && (rest[0] == Token::Subscript || rest[0] == Token::Superscript) {
            let func_run = format_run(func_text, true);
            emit_raw_with_scripts(&func_run, rest, output, &mut consumed);
            return consumed;
        }
        write_run(output, func_text, true);
        return 0;
    }

    // Symbol lookup
    if let Some(sym) = LATEX_SYMBOLS.get(cmd) {
        // Check if the symbol is followed by sub/superscript
        if !rest.is_empty() && (rest[0] == Token::Subscript || rest[0] == Token::Superscript) {
            let sym_run = format_run(sym, false);
            emit_raw_with_scripts(&sym_run, rest, output, &mut consumed);
            return consumed;
        }
        write_run(output, sym, false);
        return 0;
    }

    // Unknown command - just output the name
    write_run(output, &format!("\\{}", cmd), false);
    consumed
}

// ── Helper functions ───────────────────────────────────────────────────

fn write_run(output: &mut String, text: &str, normal: bool) {
    output.push_str("<m:r>");
    if normal {
        output.push_str("<m:rPr><m:sty m:val=\"p\"/></m:rPr>");
    }
    output.push_str("<m:t>");
    output.push_str(&xml_escape(text));
    output.push_str("</m:t></m:r>");
}

fn format_run(text: &str, normal: bool) -> String {
    let mut s = String::new();
    write_run(&mut s, text, normal);
    s
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn token_to_text(token: &Token) -> String {
    match token {
        Token::Text(t) => t.clone(),
        Token::Group(inner) => tokens_to_text(inner),
        Token::Command(cmd) => {
            if let Some(sym) = LATEX_SYMBOLS.get(cmd.as_str()) {
                sym.to_string()
            } else {
                format!("\\{}", cmd)
            }
        }
        _ => String::new(),
    }
}

fn tokens_to_text(tokens: &[Token]) -> String {
    tokens.iter().map(token_to_text).collect()
}

fn token_arg_to_omml(token: &Token, output: &mut String) {
    match token {
        Token::Group(inner) => {
            tokens_to_omml(inner, output);
        }
        Token::Text(t) => {
            write_run(output, t, false);
        }
        Token::Command(cmd) => {
            handle_command(cmd, &[], output);
        }
        _ => {}
    }
}

/// Emit text with following sub/superscript operators
fn emit_text_with_scripts(text: &str, rest: &[Token], output: &mut String, i: &mut usize) {
    let base_xml = format_run(text, false);
    emit_raw_with_scripts(&base_xml, rest, output, i);
}

/// Given raw OMML XML for a base element, handle following sub/superscript tokens
fn emit_raw_with_scripts(base_xml: &str, rest: &[Token], output: &mut String, consumed: &mut usize) {
    let mut j = 0;
    let mut sub_xml = String::new();
    let mut sup_xml = String::new();
    let mut has_sub = false;
    let mut has_sup = false;

    // Collect up to one sub and one sup in any order
    for _ in 0..2 {
        if j < rest.len() {
            if rest[j] == Token::Subscript {
                j += 1;
                has_sub = true;
                if j < rest.len() {
                    token_arg_to_omml(&rest[j], &mut sub_xml);
                    j += 1;
                }
            } else if rest[j] == Token::Superscript {
                j += 1;
                has_sup = true;
                if j < rest.len() {
                    token_arg_to_omml(&rest[j], &mut sup_xml);
                    j += 1;
                }
            } else {
                break;
            }
        }
    }

    if has_sub && has_sup {
        output.push_str(&format!(
            "<m:sSubSup><m:sSubSupPr><m:ctrlPr/></m:sSubSupPr><m:e>{}</m:e><m:sub>{}</m:sub><m:sup>{}</m:sup></m:sSubSup>",
            base_xml, sub_xml, sup_xml
        ));
    } else if has_sub {
        output.push_str(&format!(
            "<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>{}</m:e><m:sub>{}</m:sub></m:sSub>",
            base_xml, sub_xml
        ));
    } else if has_sup {
        output.push_str(&format!(
            "<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>{}</m:e><m:sup>{}</m:sup></m:sSup>",
            base_xml, sup_xml
        ));
    } else {
        output.push_str(base_xml);
    }

    // *consumed already includes the base token position advance done by caller
    // We add the script tokens consumed
    *consumed += j;
}

/// Parse \left ... \right delimiters
fn parse_left_right(tokens: &[Token]) -> (String, Vec<Token>, String, usize) {
    let mut consumed = 0;
    let mut open_delim = String::new();

    // Get opening delimiter
    if consumed < tokens.len() {
        match &tokens[consumed] {
            Token::Text(t) => {
                open_delim = t.clone();
                consumed += 1;
            }
            Token::Command(cmd) => {
                if cmd == "langle" {
                    open_delim = "\u{27E8}".to_string();
                } else if cmd == "lfloor" {
                    open_delim = "\u{230A}".to_string();
                } else if cmd == "lceil" {
                    open_delim = "\u{2308}".to_string();
                } else if cmd == "lbrace" {
                    open_delim = "{".to_string();
                } else if cmd == "vert" {
                    open_delim = "|".to_string();
                } else if cmd == "Vert" {
                    open_delim = "\u{2016}".to_string();
                } else {
                    open_delim = cmd.clone();
                }
                consumed += 1;
            }
            _ => { consumed += 1; }
        }
    }

    // Collect body until \right
    let mut body = Vec::new();
    let mut close_delim = String::new();
    let mut depth = 1; // track nested \left...\right

    while consumed < tokens.len() {
        if let Token::Command(cmd) = &tokens[consumed] {
            if cmd == "right" {
                depth -= 1;
                if depth == 0 {
                    consumed += 1;
                    // Get closing delimiter
                    if consumed < tokens.len() {
                        match &tokens[consumed] {
                            Token::Text(t) => {
                                close_delim = t.clone();
                                consumed += 1;
                            }
                            Token::Command(cmd) => {
                                if cmd == "rangle" {
                                    close_delim = "\u{27E9}".to_string();
                                } else if cmd == "rfloor" {
                                    close_delim = "\u{230B}".to_string();
                                } else if cmd == "rceil" {
                                    close_delim = "\u{2309}".to_string();
                                } else if cmd == "rbrace" {
                                    close_delim = "}".to_string();
                                } else if cmd == "vert" {
                                    close_delim = "|".to_string();
                                } else if cmd == "Vert" {
                                    close_delim = "\u{2016}".to_string();
                                } else {
                                    close_delim = cmd.clone();
                                }
                                consumed += 1;
                            }
                            _ => { consumed += 1; }
                        }
                    }
                    break;
                }
                body.push(tokens[consumed].clone());
                consumed += 1;
                continue;
            } else if cmd == "left" {
                depth += 1;
            }
        }
        body.push(tokens[consumed].clone());
        consumed += 1;
    }

    // Handle "." as invisible delimiter
    if open_delim == "." { open_delim = String::new(); }
    if close_delim == "." { close_delim = String::new(); }

    (open_delim, body, close_delim, consumed)
}

/// Find body tokens of an environment (everything between \begin{env} and \end{env})
fn find_env_body(tokens: &[Token], env_name: &str) -> (Vec<Token>, usize) {
    let mut body = Vec::new();
    let mut consumed = 0;
    let mut depth = 1;

    while consumed < tokens.len() {
        if let Token::Command(cmd) = &tokens[consumed] {
            if cmd == "begin" && consumed + 1 < tokens.len() {
                if let Token::Group(inner) = &tokens[consumed + 1] {
                    if tokens_to_text(inner) == env_name {
                        depth += 1;
                    }
                }
                body.push(tokens[consumed].clone());
                consumed += 1;
                body.push(tokens[consumed].clone());
                consumed += 1;
                continue;
            }
            if cmd == "end" && consumed + 1 < tokens.len() {
                if let Token::Group(inner) = &tokens[consumed + 1] {
                    if tokens_to_text(inner) == env_name {
                        depth -= 1;
                        if depth == 0 {
                            consumed += 2; // skip \end{env}
                            break;
                        }
                    }
                }
                body.push(tokens[consumed].clone());
                consumed += 1;
                body.push(tokens[consumed].clone());
                consumed += 1;
                continue;
            }
        }
        body.push(tokens[consumed].clone());
        consumed += 1;
    }

    (body, consumed)
}

/// Parse matrix body into rows of cells (split by \\ and &)
fn parse_matrix_body(tokens: &[Token]) -> Vec<Vec<Vec<Token>>> {
    let mut rows: Vec<Vec<Vec<Token>>> = Vec::new();
    let mut current_row: Vec<Vec<Token>> = Vec::new();
    let mut current_cell: Vec<Token> = Vec::new();

    for token in tokens {
        match token {
            Token::Newline => {
                current_row.push(current_cell);
                current_cell = Vec::new();
                rows.push(current_row);
                current_row = Vec::new();
            }
            Token::Ampersand => {
                current_row.push(current_cell);
                current_cell = Vec::new();
            }
            _ => {
                current_cell.push(token.clone());
            }
        }
    }

    // Don't forget last cell/row
    if !current_cell.is_empty() || !current_row.is_empty() {
        current_row.push(current_cell);
        rows.push(current_row);
    }

    rows
}

/// Emit a matrix with optional delimiters
fn emit_matrix(output: &mut String, rows: &[Vec<Vec<Token>>], beg_chr: &str, end_chr: &str) {
    let ncols = rows.iter().map(|r| r.len()).max().unwrap_or(0);

    if !beg_chr.is_empty() || !end_chr.is_empty() {
        output.push_str("<m:d>");
        output.push_str("<m:dPr>");
        output.push_str(&format!("<m:begChr m:val=\"{}\"/>", xml_escape(beg_chr)));
        output.push_str(&format!("<m:endChr m:val=\"{}\"/>", xml_escape(end_chr)));
        output.push_str("<m:ctrlPr/></m:dPr>");
        output.push_str("<m:e>");
    }

    output.push_str("<m:m>");
    output.push_str("<m:mPr>");
    output.push_str(&format!("<m:mcs><m:mc><m:mcPr><m:count m:val=\"{}\"/><m:mcJc m:val=\"center\"/></m:mcPr></m:mc></m:mcs>", ncols));
    output.push_str("<m:ctrlPr/></m:mPr>");

    for row in rows {
        output.push_str("<m:mr>");
        for cell in row {
            output.push_str("<m:e>");
            if cell.is_empty() {
                // Empty cell
            } else {
                tokens_to_omml(cell, output);
            }
            output.push_str("</m:e>");
        }
        // Pad missing cells
        for _ in row.len()..ncols {
            output.push_str("<m:e/>");
        }
        output.push_str("</m:mr>");
    }
    output.push_str("</m:m>");

    if !beg_chr.is_empty() || !end_chr.is_empty() {
        output.push_str("</m:e></m:d>");
    }
}

/// Emit an equation array (align environment)
fn emit_equation_array(output: &mut String, rows: &[Vec<Vec<Token>>]) {
    output.push_str("<m:eqArr>");
    output.push_str("<m:eqArrPr><m:ctrlPr/></m:eqArrPr>");

    for row in rows {
        output.push_str("<m:e>");
        for (j, cell) in row.iter().enumerate() {
            if j > 0 {
                // Alignment point
                write_run(output, "=", false);
            }
            tokens_to_omml(cell, output);
        }
        output.push_str("</m:e>");
    }

    output.push_str("</m:eqArr>");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_expression() {
        let omml = latex_to_omml("x + y");
        assert!(omml.contains("<m:r>"));
        assert!(omml.contains("x"));
        assert!(omml.contains("+"));
        assert!(omml.contains("y"));
    }

    #[test]
    fn test_fraction() {
        let omml = latex_to_omml("\\frac{a}{b}");
        assert!(omml.contains("<m:f>"));
        assert!(omml.contains("<m:num>"));
        assert!(omml.contains("<m:den>"));
    }

    #[test]
    fn test_superscript() {
        let omml = latex_to_omml("x^{2}");
        assert!(omml.contains("<m:sSup>"));
        assert!(omml.contains("<m:sup>"));
    }

    #[test]
    fn test_subscript() {
        let omml = latex_to_omml("x_{i}");
        assert!(omml.contains("<m:sSub>"));
        assert!(omml.contains("<m:sub>"));
    }

    #[test]
    fn test_sqrt() {
        let omml = latex_to_omml("\\sqrt{x}");
        assert!(omml.contains("<m:rad>"));
        assert!(omml.contains("<m:degHide"));
    }

    #[test]
    fn test_greek() {
        let omml = latex_to_omml("\\alpha + \\beta");
        assert!(omml.contains("\u{03B1}"));
        assert!(omml.contains("\u{03B2}"));
    }

    #[test]
    fn test_sum_with_limits() {
        let omml = latex_to_omml("\\sum_{i=1}^{n}");
        assert!(omml.contains("<m:nary>"));
        assert!(omml.contains("\u{2211}"));
        assert!(omml.contains("<m:sub>"));
        assert!(omml.contains("<m:sup>"));
    }

    #[test]
    fn test_display_math_paragraph() {
        let omml = latex_to_omml_paragraph("E = mc^{2}");
        assert!(omml.contains("<m:oMathPara>"));
        assert!(omml.contains("<m:oMath>"));
        assert!(omml.contains("<m:sSup>"));
    }

    #[test]
    fn test_inline_math() {
        let omml = latex_to_omml_inline("x^2");
        assert!(omml.contains("<m:oMath>"));
        assert!(!omml.contains("<m:oMathPara>"));
    }
}
