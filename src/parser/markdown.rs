//! Markdown parser using pulldown-cmark
//!
//! Converts raw markdown text into our AST types defined in `ast.rs`.

use crate::parser::ast::*;
use once_cell::sync::Lazy;
use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd};
use regex::Regex;
use std::collections::HashMap;

// Include patterns - match whole line directives
static INCLUDE_PATTERN: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r"^\{!include:([^}]+)\}$").expect("INCLUDE_PATTERN regex should be valid")
});

static CODE_INCLUDE_PATTERN: Lazy<Regex> = Lazy::new(|| {
    // Matches: {!code:path} or {!code:path:start-end} or {!code:path:start-end:lang}
    Regex::new(r"^\{!code:([^:}]+)(?::(\d+)-(\d+))?(?::([a-zA-Z0-9]+))?\}$")
        .expect("CODE_INCLUDE_PATTERN regex should be valid")
});

static HTML_ID_PATTERN: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r"<!--\s*\{#([a-zA-Z0-9_:-]+)\}\s*-->")
        .expect("HTML_ID_PATTERN regex should be valid")
});

static TABLE_CAPTION_PATTERN: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r"^Table:\s*(.*)\s*\{#([a-zA-Z0-9_:-]+)\}$")
        .expect("TABLE_CAPTION_PATTERN regex should be valid")
});

static TABLE_CAPTION_NO_ID_PATTERN: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r"^Table:\s*(.*)$").expect("TABLE_CAPTION_NO_ID_PATTERN regex should be valid")
});

/// Builder for footnote definitions
struct FootnoteBuilder {
    name: String,
    content: Vec<Block>,
}

/// Parse markdown text into a ParsedDocument
pub fn parse_markdown(input: &str) -> ParsedDocument {
    let parser = Parser::new_ext(input, get_parser_options());

    let mut blocks = Vec::new();
    let mut footnotes = HashMap::new();
    let mut current_block: Option<BlockBuilder> = None;
    let mut block_stack: Vec<BlockBuilder> = Vec::new();
    let mut list_stack: Vec<ListBuilder> = Vec::new();
    let mut list_item_inlines: Vec<Inline> = Vec::new(); // Track inlines for current list item
    let mut table_builder: Option<TableBuilder> = None;
    let mut footnote_builder: Option<FootnoteBuilder> = None;

    // Inline element stack for handling nested formatting
    let mut inline_stack: Vec<InlineBuilder> = Vec::new();
    let mut current_inlines: Vec<Inline> = Vec::new();

    for event in parser {
        match event {
            // Block-level events
            Event::Start(tag) => {
                match tag {
                    Tag::Heading { level, .. } => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        current_block = Some(BlockBuilder::Heading {
                            level: level as u8,
                            content: Vec::new(),
                            id: None,
                        });
                        current_inlines = Vec::new();
                    }
                    Tag::Paragraph => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        current_block = Some(BlockBuilder::Paragraph(Vec::new()));
                        current_inlines = Vec::new();
                    }
                    Tag::BlockQuote(_) => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        block_stack.push(BlockBuilder::BlockQuote(Vec::new()));
                    }
                    Tag::List(start_number) => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        let ordered = start_number.is_some();
                        list_stack.push(ListBuilder {
                            ordered,
                            start: start_number.map(|n| n as u32),
                            items: Vec::new(),
                        });
                    }
                    Tag::Item => {
                        // Finish any previous item's inlines
                        if let Some(list) = list_stack.last_mut() {
                            if let Some(item) = list.items.last_mut() {
                                if !list_item_inlines.is_empty() {
                                    item.content
                                        .push(Block::Paragraph(list_item_inlines.clone()));
                                }
                            }
                            list.items.push(ListItemBuilder {
                                content: Vec::new(),
                                checked: None,
                            });
                        }
                        list_item_inlines = Vec::new();
                    }
                    Tag::CodeBlock(kind) => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        let info = match kind {
                            pulldown_cmark::CodeBlockKind::Fenced(info) => info.to_string(),
                            pulldown_cmark::CodeBlockKind::Indented => String::new(),
                        };
                        let (lang, filename, highlight_lines, show_line_numbers) =
                            parse_code_block_info(&info);
                        current_block = Some(BlockBuilder::CodeBlock {
                            lang,
                            content: String::new(),
                            filename,
                            highlight_lines,
                            show_line_numbers,
                        });
                    }
                    Tag::Table(alignment) => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        table_builder = Some(TableBuilder {
                            alignments: alignment
                                .into_iter()
                                .map(|a| match a {
                                    pulldown_cmark::Alignment::Left => Alignment::Left,
                                    pulldown_cmark::Alignment::Center => Alignment::Center,
                                    pulldown_cmark::Alignment::Right => Alignment::Right,
                                    pulldown_cmark::Alignment::None => Alignment::None,
                                })
                                .collect(),
                            headers: Vec::new(),
                            rows: Vec::new(),
                            current_row: Vec::new(),
                            current_cell: Vec::new(),
                        });
                    }
                    Tag::TableHead => {
                        if let Some(table) = table_builder.as_mut() {
                            table.current_row = Vec::new();
                        }
                    }
                    Tag::TableRow => {
                        if let Some(table) = table_builder.as_mut() {
                            table.current_row = Vec::new();
                        }
                    }
                    Tag::TableCell => {
                        if let Some(table) = table_builder.as_mut() {
                            table.current_cell = Vec::new();
                        }
                    }
                    Tag::FootnoteDefinition(name) => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        footnote_builder = Some(FootnoteBuilder {
                            name: name.to_string(),
                            content: Vec::new(),
                        });
                        current_inlines = Vec::new();
                    }
                    // Inline formatting - push to stack
                    Tag::Emphasis => {
                        inline_stack.push(InlineBuilder::Italic(Vec::new()));
                    }
                    Tag::Strong => {
                        inline_stack.push(InlineBuilder::Bold(Vec::new()));
                    }
                    Tag::Strikethrough => {
                        inline_stack.push(InlineBuilder::Strikethrough(Vec::new()));
                    }
                    Tag::Link {
                        dest_url, title, ..
                    } => {
                        inline_stack.push(InlineBuilder::Link {
                            text: Vec::new(),
                            url: dest_url.to_string(),
                            title: if title.is_empty() {
                                None
                            } else {
                                Some(title.to_string())
                            },
                        });
                    }
                    Tag::Image {
                        dest_url, title, ..
                    } => {
                        inline_stack.push(InlineBuilder::Image {
                            alt: String::new(),
                            src: dest_url.to_string(),
                            title: if title.is_empty() {
                                None
                            } else {
                                Some(title.to_string())
                            },
                        });
                    }
                    _ => {}
                }
            }
            Event::End(tag_end) => {
                match tag_end {
                    TagEnd::Heading(_) => {
                        if let Some(BlockBuilder::Heading { level, id, .. }) = current_block.take()
                        {
                            let (content, id) = extract_anchor_id(current_inlines, id);
                            blocks.push(Block::Heading { level, content, id });
                        }
                        current_inlines = Vec::new();
                    }
                    TagEnd::Paragraph => {
                        if !block_stack.is_empty() {
                            // Handle paragraph end inside blockquotes
                            if let Some(BlockBuilder::BlockQuote(content)) = block_stack.last_mut()
                            {
                                if !current_inlines.is_empty() {
                                    content.push(Block::Paragraph(current_inlines.clone()));
                                }
                            }
                        } else if footnote_builder.is_some() {
                            // Handle paragraph end inside footnote definitions
                            if let Some(builder) = footnote_builder.as_mut() {
                                if !current_inlines.is_empty() {
                                    builder
                                        .content
                                        .push(Block::Paragraph(current_inlines.clone()));
                                }
                            }
                        } else if let Some(BlockBuilder::Paragraph(_)) = current_block.take() {
                            // Check for single image or image + attributes
                            // Case 1: Paragraph contains only an image
                            // Case 2: Paragraph contains image + attributes text (e.g. {width=50%})
                            let is_image_block = if current_inlines.len() == 1 {
                                matches!(current_inlines[0], Inline::Image { .. })
                            } else if current_inlines.len() == 2 {
                                matches!(current_inlines[0], Inline::Image { .. })
                                    && matches!(current_inlines[1], Inline::Text(ref t) if extract_image_attributes(t).is_some())
                            } else {
                                false
                            };

                            if is_image_block {
                                let (image, width) = if current_inlines.len() == 1 {
                                    (current_inlines.remove(0), None)
                                } else {
                                    let attrs = current_inlines
                                        .pop()
                                        .expect("attrs should exist when len == 2"); // Text
                                    let img = current_inlines.remove(0); // Image
                                    let width = if let Inline::Text(t) = attrs {
                                        extract_image_attributes(&t)
                                    } else {
                                        None
                                    };
                                    (img, width)
                                };

                                if let Inline::Image { alt, src, title } = image {
                                    add_block_to_correct_stack(
                                        &mut blocks,
                                        &mut footnote_builder,
                                        &mut list_stack,
                                        &mut block_stack,
                                        Block::Image {
                                            alt,
                                            src,
                                            title,
                                            width,
                                            id: None,
                                        },
                                    );
                                    current_inlines = Vec::new();
                                    continue;
                                }
                            }
                            add_block_to_correct_stack(
                                &mut blocks,
                                &mut footnote_builder,
                                &mut list_stack,
                                &mut block_stack,
                                Block::Paragraph(current_inlines.clone()),
                            );
                        }
                        current_inlines = Vec::new();
                    }
                    TagEnd::BlockQuote(_) => {
                        if let Some(BlockBuilder::BlockQuote(content)) = block_stack.pop() {
                            add_block_to_correct_stack(
                                &mut blocks,
                                &mut footnote_builder,
                                &mut list_stack,
                                &mut block_stack,
                                Block::BlockQuote(content),
                            );
                        }
                    }
                    TagEnd::List(_) => {
                        if let Some(list) = list_stack.pop() {
                            add_block_to_correct_stack(
                                &mut blocks,
                                &mut footnote_builder,
                                &mut list_stack,
                                &mut block_stack,
                                Block::List {
                                    ordered: list.ordered,
                                    start: list.start,
                                    items: list
                                        .items
                                        .into_iter()
                                        .map(|item| ListItem {
                                            content: item.content,
                                            checked: item.checked,
                                        })
                                        .collect(),
                                },
                            );
                        }
                    }
                    TagEnd::Item => {
                        // Finish current item's inlines
                        if let Some(list) = list_stack.last_mut() {
                            if let Some(item) = list.items.last_mut() {
                                if !list_item_inlines.is_empty() {
                                    item.content
                                        .push(Block::Paragraph(list_item_inlines.clone()));
                                    list_item_inlines = Vec::new();
                                }
                            }
                        }
                    }
                    TagEnd::CodeBlock => {
                        if let Some(BlockBuilder::CodeBlock {
                            lang,
                            content,
                            filename,
                            highlight_lines,
                            show_line_numbers,
                        }) = current_block.take()
                        {
                            let block = if lang.as_deref() == Some("mermaid") {
                                Block::Mermaid { content, id: None }
                            } else {
                                Block::CodeBlock {
                                    lang,
                                    content,
                                    filename,
                                    highlight_lines,
                                    show_line_numbers,
                                }
                            };
                            add_block_to_correct_stack(
                                &mut blocks,
                                &mut footnote_builder,
                                &mut list_stack,
                                &mut block_stack,
                                block,
                            );
                        }
                    }
                    TagEnd::Table => {
                        if let Some(table) = table_builder.take() {
                            let mut caption = None;
                            let mut id = None;

                            // Check if the preceding block was an HTML comment with an ID
                            // or a paragraph that looks like a table caption.
                            // We need to check the correct stack based on context.
                            let last_block = if let Some(builder) = footnote_builder.as_mut() {
                                builder.content.last_mut()
                            } else if let Some(list) = list_stack.last_mut() {
                                if let Some(item) = list.items.last_mut() {
                                    item.content.last_mut()
                                } else {
                                    None
                                }
                            } else if let Some(BlockBuilder::BlockQuote(content)) =
                                block_stack.last_mut()
                            {
                                content.last_mut()
                            } else {
                                blocks.last_mut()
                            };

                            if let Some(block) = last_block {
                                match block {
                                    Block::Html(html) => {
                                        if let Some(cap) = HTML_ID_PATTERN.captures(html) {
                                            id = Some(
                                                cap.get(1)
                                                    .expect("HTML_ID_PATTERN should have capture group 1")
                                                    .as_str()
                                                    .to_string(),
                                            );
                                            // Mark for removal by changing to something else or we'll pop it
                                        }
                                    }
                                    Block::Paragraph(inlines) => {
                                        let text = extract_inline_text(inlines);
                                        if let Some(cap) = TABLE_CAPTION_PATTERN.captures(&text) {
                                            caption = Some(
                                                cap.get(1)
                                                    .expect("TABLE_CAPTION_PATTERN should have capture group 1")
                                                    .as_str()
                                                    .trim()
                                                    .to_string(),
                                            );
                                            id = Some(
                                                cap.get(2)
                                                    .expect("TABLE_CAPTION_PATTERN should have capture group 2")
                                                    .as_str()
                                                    .to_string(),
                                            );
                                        } else if let Some(cap) =
                                            TABLE_CAPTION_NO_ID_PATTERN.captures(&text)
                                        {
                                            caption = Some(
                                                cap.get(1)
                                                    .expect("TABLE_CAPTION_NO_ID_PATTERN should have capture group 1")
                                                    .as_str()
                                                    .trim()
                                                    .to_string(),
                                            );
                                        }
                                    }
                                    _ => {}
                                }
                            }

                            // If we found a caption or ID, we need to remove that last block
                            if id.is_some() || caption.is_some() {
                                if let Some(builder) = footnote_builder.as_mut() {
                                    builder.content.pop();
                                } else if let Some(list) = list_stack.last_mut() {
                                    if let Some(item) = list.items.last_mut() {
                                        item.content.pop();
                                    }
                                } else if let Some(BlockBuilder::BlockQuote(content)) =
                                    block_stack.last_mut()
                                {
                                    content.pop();
                                } else {
                                    blocks.pop();
                                }
                            }

                            add_block_to_correct_stack(
                                &mut blocks,
                                &mut footnote_builder,
                                &mut list_stack,
                                &mut block_stack,
                                Block::Table {
                                    headers: table.headers,
                                    alignments: table.alignments,
                                    rows: table.rows,
                                    caption,
                                    id,
                                },
                            );
                        }
                    }
                    TagEnd::TableHead => {
                        if let Some(table) = table_builder.as_mut() {
                            table.headers = table
                                .current_row
                                .drain(..)
                                .map(|mut cell| {
                                    cell.is_header = true;
                                    cell
                                })
                                .collect();
                        }
                    }
                    TagEnd::TableRow => {
                        if let Some(table) = table_builder.as_mut() {
                            table.rows.push(table.current_row.drain(..).collect());
                        }
                    }
                    TagEnd::TableCell => {
                        if let Some(table) = table_builder.as_mut() {
                            table.current_row.push(TableCell {
                                content: table.current_cell.drain(..).collect(),
                                is_header: false,
                            });
                        }
                    }
                    TagEnd::Emphasis => {
                        if let Some(InlineBuilder::Italic(content)) = inline_stack.pop() {
                            if !list_stack.is_empty() {
                                list_item_inlines.push(Inline::Italic(content));
                            } else {
                                add_inline(&mut current_inlines, Inline::Italic(content));
                            }
                        }
                    }
                    TagEnd::Strong => {
                        if let Some(InlineBuilder::Bold(content)) = inline_stack.pop() {
                            if !list_stack.is_empty() {
                                list_item_inlines.push(Inline::Bold(content));
                            } else {
                                add_inline(&mut current_inlines, Inline::Bold(content));
                            }
                        }
                    }
                    TagEnd::Strikethrough => {
                        if let Some(InlineBuilder::Strikethrough(content)) = inline_stack.pop() {
                            if !list_stack.is_empty() {
                                list_item_inlines.push(Inline::Strikethrough(content));
                            } else {
                                add_inline(&mut current_inlines, Inline::Strikethrough(content));
                            }
                        }
                    }
                    TagEnd::Link => {
                        if let Some(InlineBuilder::Link { text, url, title }) = inline_stack.pop() {
                            if !list_stack.is_empty() {
                                list_item_inlines.push(Inline::Link { text, url, title });
                            } else {
                                add_inline(&mut current_inlines, Inline::Link { text, url, title });
                            }
                        }
                    }
                    TagEnd::Image => {
                        if let Some(InlineBuilder::Image { alt, src, title }) = inline_stack.pop() {
                            if !list_stack.is_empty() {
                                list_item_inlines.push(Inline::Image { alt, src, title });
                            } else {
                                add_inline(&mut current_inlines, Inline::Image { alt, src, title });
                            }
                        }
                    }
                    TagEnd::FootnoteDefinition => {
                        finish_current_block_with_footnote(
                            &mut current_block,
                            &mut blocks,
                            &mut footnote_builder,
                            &mut list_stack,
                            &mut block_stack,
                        );
                        if let Some(builder) = footnote_builder.take() {
                            footnotes.insert(builder.name, builder.content);
                        }
                        current_inlines = Vec::new();
                    }
                    _ => {}
                }
            }
            Event::Text(text) => {
                let text = text.to_string();
                if let Some(table) = table_builder.as_mut() {
                    table.current_cell.push(Inline::Text(text));
                } else if !list_stack.is_empty() {
                    // Handle text inside list items - accumulate into list_item_inlines
                    if inline_stack.is_empty() {
                        list_item_inlines.push(Inline::Text(text));
                    } else {
                        add_text_to_inline_stack(&mut inline_stack, &mut list_item_inlines, text);
                    }
                } else if let Some(block) = current_block.as_mut() {
                    match block {
                        BlockBuilder::Heading { .. } | BlockBuilder::Paragraph(_) => {
                            add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, text);
                        }
                        BlockBuilder::CodeBlock { content, .. } => {
                            content.push_str(&text);
                        }
                        _ => {}
                    }
                } else if footnote_builder.is_some() {
                    // Handle text inside footnote definitions
                    add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, text);
                } else if !block_stack.is_empty() {
                    // Handle text inside blockquotes
                    add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, text);
                }
            }
            Event::Code(code) => {
                let code = code.to_string();
                if let Some(table) = table_builder.as_mut() {
                    table.current_cell.push(Inline::Code(code));
                } else if !list_stack.is_empty() {
                    // Handle code inside list items
                    if inline_stack.is_empty() {
                        list_item_inlines.push(Inline::Code(code));
                    } else {
                        add_text_to_inline_stack(
                            &mut inline_stack,
                            &mut list_item_inlines,
                            code.clone(),
                        );
                        // Convert to code
                        if let Some(inline) = list_item_inlines.last_mut() {
                            if let Inline::Text(t) = inline {
                                *inline = Inline::Code(t.clone());
                            }
                        }
                    }
                } else if let Some(block) = current_block.as_mut() {
                    match block {
                        BlockBuilder::Heading { .. } | BlockBuilder::Paragraph(_) => {
                            add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, code);
                            // Convert the last text to code
                            if let Some(inline) = current_inlines.last_mut() {
                                if let Inline::Text(text) = inline {
                                    *inline = Inline::Code(text.clone());
                                }
                            }
                        }
                        _ => {}
                    }
                } else if footnote_builder.is_some() {
                    // Handle code inside footnote definitions
                    add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, code);
                    if let Some(inline) = current_inlines.last_mut() {
                        if let Inline::Text(text) = inline {
                            *inline = Inline::Code(text.clone());
                        }
                    }
                } else if !block_stack.is_empty() {
                    // Handle code inside blockquotes
                    add_text_to_inline_stack(&mut inline_stack, &mut current_inlines, code);
                    if let Some(inline) = current_inlines.last_mut() {
                        if let Inline::Text(text) = inline {
                            *inline = Inline::Code(text.clone());
                        }
                    }
                }
            }

            Event::SoftBreak => {
                if let Some(table) = table_builder.as_mut() {
                    table.current_cell.push(Inline::SoftBreak);
                } else if !list_stack.is_empty() {
                    // Soft break becomes a space in list items
                    list_item_inlines.push(Inline::Text(" ".to_string()));
                } else if let Some(block) = current_block.as_mut() {
                    match block {
                        BlockBuilder::Heading { .. } | BlockBuilder::Paragraph(_) => {
                            // Soft break becomes a space (or ignored for Thai if we implement it)
                            // For now, Word handles wrapping, so we convert soft break to space
                            // to prevent words from sticking together, unless it's Thai text?
                            // Actually, standard markdown behavior is soft break -> space or newline.
                            // Let's stick to space for now to match CommonMark.
                            add_text_to_inline_stack(
                                &mut inline_stack,
                                &mut current_inlines,
                                " ".to_string(),
                            );
                        }
                        _ => {}
                    }
                } else if footnote_builder.is_some() {
                    add_text_to_inline_stack(
                        &mut inline_stack,
                        &mut current_inlines,
                        " ".to_string(),
                    );
                } else if !block_stack.is_empty() {
                    add_inline(&mut current_inlines, Inline::SoftBreak);
                }
            }
            Event::HardBreak => {
                if let Some(table) = table_builder.as_mut() {
                    table.current_cell.push(Inline::HardBreak);
                } else if !list_stack.is_empty() {
                    // Hard break in list items
                    list_item_inlines.push(Inline::HardBreak);
                } else if let Some(block) = current_block.as_mut() {
                    match block {
                        BlockBuilder::Heading { .. } | BlockBuilder::Paragraph(_) => {
                            add_inline(&mut current_inlines, Inline::HardBreak);
                        }
                        _ => {}
                    }
                } else if footnote_builder.is_some() {
                    // Handle hard breaks inside footnote definitions
                    add_inline(&mut current_inlines, Inline::HardBreak);
                } else if !block_stack.is_empty() {
                    // Handle hard breaks inside blockquotes
                    add_inline(&mut current_inlines, Inline::HardBreak);
                }
            }
            Event::Rule => {
                finish_current_block_with_footnote(
                    &mut current_block,
                    &mut blocks,
                    &mut footnote_builder,
                    &mut list_stack,
                    &mut block_stack,
                );
                add_block_to_correct_stack(
                    &mut blocks,
                    &mut footnote_builder,
                    &mut list_stack,
                    &mut block_stack,
                    Block::ThematicBreak,
                );
            }
            Event::Html(html) => {
                finish_current_block_with_footnote(
                    &mut current_block,
                    &mut blocks,
                    &mut footnote_builder,
                    &mut list_stack,
                    &mut block_stack,
                );
                add_block_to_correct_stack(
                    &mut blocks,
                    &mut footnote_builder,
                    &mut list_stack,
                    &mut block_stack,
                    Block::Html(html.to_string()),
                );
            }
            Event::FootnoteReference(name) => {
                let name = name.to_string();
                if let Some(table) = table_builder.as_mut() {
                    table.current_cell.push(Inline::FootnoteRef(name));
                } else if !list_stack.is_empty() {
                    // Footnote reference inside list items
                    list_item_inlines.push(Inline::FootnoteRef(name));
                } else if let Some(block) = current_block.as_mut() {
                    match block {
                        BlockBuilder::Heading { .. } | BlockBuilder::Paragraph(_) => {
                            add_inline(&mut current_inlines, Inline::FootnoteRef(name));
                        }
                        _ => {}
                    }
                } else if let Some(BlockBuilder::BlockQuote(content)) = block_stack.last_mut() {
                    // Handle footnote references inside blockquotes
                    if let Some(Block::Paragraph(inlines)) = content.last_mut() {
                        inlines.push(Inline::FootnoteRef(name));
                    }
                }
            }
            Event::TaskListMarker(checked) => {
                if let Some(list) = list_stack.last_mut() {
                    if let Some(item) = list.items.last_mut() {
                        item.checked = Some(checked);
                    }
                }
            }
            _ => {}
        }
    }

    // Don't forget the last block
    finish_current_block_with_footnote(
        &mut current_block,
        &mut blocks,
        &mut footnote_builder,
        &mut list_stack,
        &mut block_stack,
    );

    // Process cross-references
    let blocks = process_blocks_for_cross_refs(blocks);

    // Process include directives
    let blocks = process_include_directives(blocks);

    ParsedDocument {
        frontmatter: None,
        blocks,
        footnotes,
    }
}

/// Finish the current block and add it to blocks
#[allow(dead_code)]
fn finish_current_block(current_block: &mut Option<BlockBuilder>, blocks: &mut Vec<Block>) {
    if let Some(block) = current_block.take() {
        blocks.push(block.build());
    }
}

/// Finish the current block and add it to blocks or footnote content
fn finish_current_block_with_footnote(
    current_block: &mut Option<BlockBuilder>,
    blocks: &mut Vec<Block>,
    footnote_builder: &mut Option<FootnoteBuilder>,
    list_stack: &mut [ListBuilder],
    block_stack: &mut [BlockBuilder],
) {
    if let Some(builder) = current_block.take() {
        let block = builder.build();
        add_block_to_correct_stack(blocks, footnote_builder, list_stack, block_stack, block);
    }
}

/// Add a block to the correct stack (footnote, list, blockquote, or top-level)
fn add_block_to_correct_stack(
    blocks: &mut Vec<Block>,
    footnote_builder: &mut Option<FootnoteBuilder>,
    list_stack: &mut [ListBuilder],
    block_stack: &mut [BlockBuilder],
    block: Block,
) {
    if let Some(builder) = footnote_builder {
        builder.content.push(block);
    } else if let Some(list) = list_stack.last_mut() {
        if let Some(item) = list.items.last_mut() {
            item.content.push(block);
        } else {
            blocks.push(block);
        }
    } else if let Some(BlockBuilder::BlockQuote(content)) = block_stack.last_mut() {
        content.push(block);
    } else {
        blocks.push(block);
    }
}

/// Process blocks to extract cross-references from inline content
fn process_blocks_for_cross_refs(blocks: Vec<Block>) -> Vec<Block> {
    blocks
        .into_iter()
        .map(|block| match block {
            Block::Paragraph(inlines) => Block::Paragraph(process_cross_refs(inlines)),
            Block::Heading { level, content, id } => Block::Heading {
                level,
                content: process_cross_refs(content),
                id,
            },
            Block::Table {
                headers,
                alignments,
                rows,
                caption,
                id,
            } => Block::Table {
                headers: headers
                    .into_iter()
                    .map(|c| TableCell {
                        content: process_cross_refs(c.content),
                        is_header: c.is_header,
                    })
                    .collect(),
                alignments,
                rows: rows
                    .into_iter()
                    .map(|r| {
                        r.into_iter()
                            .map(|c| TableCell {
                                content: process_cross_refs(c.content),
                                is_header: c.is_header,
                            })
                            .collect()
                    })
                    .collect(),
                caption,
                id,
            },
            Block::BlockQuote(inner) => Block::BlockQuote(process_blocks_for_cross_refs(inner)),
            Block::List {
                ordered,
                start,
                items,
            } => Block::List {
                ordered,
                start,
                items: items
                    .into_iter()
                    .map(|item| ListItem {
                        content: process_blocks_for_cross_refs(item.content),
                        checked: item.checked,
                    })
                    .collect(),
            },
            // Tables, CodeBlocks, Images, etc. - leave as is
            other => other,
        })
        .collect()
}

/// Process blocks to detect include directives
fn process_include_directives(blocks: Vec<Block>) -> Vec<Block> {
    blocks
        .into_iter()
        .flat_map(|block| {
            match block {
                Block::Paragraph(ref inlines) => {
                    // Check if this is a single-text paragraph that's an include directive
                    if inlines.len() == 1 {
                        if let Inline::Text(text) = &inlines[0] {
                            let text = text.trim();

                            // Check for {!include:...}
                            if let Some(cap) = INCLUDE_PATTERN.captures(text) {
                                let path = cap
                                    .get(1)
                                    .expect("INCLUDE_PATTERN should have capture group 1")
                                    .as_str()
                                    .to_string();
                                return vec![Block::Include {
                                    path,
                                    resolved: None,
                                }];
                            }

                            // Check for {!code:...}
                            if let Some(cap) = CODE_INCLUDE_PATTERN.captures(text) {
                                let path = cap
                                    .get(1)
                                    .expect("CODE_INCLUDE_PATTERN should have capture group 1")
                                    .as_str()
                                    .to_string();
                                let start_line = cap.get(2).map(|m| {
                                    m.as_str()
                                        .parse::<u32>()
                                        .expect("start_line should be valid u32")
                                });
                                let end_line = cap.get(3).map(|m| {
                                    m.as_str()
                                        .parse::<u32>()
                                        .expect("end_line should be valid u32")
                                });
                                let lang = cap.get(4).map(|m| m.as_str().to_string());

                                return vec![Block::CodeInclude {
                                    path,
                                    start_line,
                                    end_line,
                                    lang,
                                }];
                            }
                        }
                    }
                    vec![block]
                }
                // Recursively process blockquotes and lists
                Block::BlockQuote(inner) => {
                    vec![Block::BlockQuote(process_include_directives(inner))]
                }
                Block::List {
                    ordered,
                    start,
                    items,
                } => {
                    let processed_items = items
                        .into_iter()
                        .map(|item| ListItem {
                            content: process_include_directives(item.content),
                            checked: item.checked,
                        })
                        .collect();
                    vec![Block::List {
                        ordered,
                        start,
                        items: processed_items,
                    }]
                }
                other => vec![other],
            }
        })
        .collect()
}

/// Process inlines to extract cross-references from text
/// Converts `{ref:target}` patterns in text to Inline::CrossRef
fn process_cross_refs(inlines: Vec<Inline>) -> Vec<Inline> {
    let cross_ref_pattern = regex::Regex::new(r"\{ref:([a-zA-Z0-9_:-]+)\}")
        .expect("cross_ref_pattern regex should be valid");

    let mut result = Vec::new();

    for inline in inlines {
        match inline {
            Inline::Text(text) => {
                let mut last_end = 0;

                for cap in cross_ref_pattern.captures_iter(&text) {
                    let match_start = cap
                        .get(0)
                        .expect("cross_ref_pattern should have capture group 0")
                        .start();
                    let match_end = cap
                        .get(0)
                        .expect("cross_ref_pattern should have capture group 0")
                        .end();

                    // Add text before the match
                    if match_start > last_end {
                        result.push(Inline::Text(text[last_end..match_start].to_string()));
                    }

                    // Parse the reference target
                    let target = cap
                        .get(1)
                        .expect("cross_ref_pattern should have capture group 1")
                        .as_str();
                    let (ref_type, actual_target) = parse_ref_target(target);

                    result.push(Inline::CrossRef {
                        target: actual_target.to_string(),
                        ref_type,
                    });

                    last_end = match_end;
                }

                // Add remaining text after last match (or all text if no matches)
                if last_end < text.len() {
                    result.push(Inline::Text(text[last_end..].to_string()));
                }
            }
            // Recursively process nested inlines
            Inline::Bold(inner) => {
                result.push(Inline::Bold(process_cross_refs(inner)));
            }
            Inline::Italic(inner) => {
                result.push(Inline::Italic(process_cross_refs(inner)));
            }
            Inline::Strikethrough(inner) => {
                result.push(Inline::Strikethrough(inner));
            }
            Inline::Link { text, url, title } => {
                result.push(Inline::Link {
                    text: process_cross_refs(text),
                    url,
                    title,
                });
            }
            // Keep other inlines as-is
            other => result.push(other),
        }
    }

    result
}

/// Parse reference target to extract type prefix
/// "fig:diagram" -> (RefType::Figure, "diagram")
/// "intro" -> (RefType::Unknown, "intro")
fn parse_ref_target(target: &str) -> (RefType, &str) {
    if let Some(colon_pos) = target.find(':') {
        let prefix = &target[..colon_pos];
        let id = &target[colon_pos + 1..];
        (RefType::from_prefix(prefix), id)
    } else {
        (RefType::Unknown, target)
    }
}

/// Add text to the appropriate place in the inline stack
fn add_text_to_inline_stack(
    inline_stack: &mut [InlineBuilder],
    current_inlines: &mut Vec<Inline>,
    text: String,
) {
    if let Some(builder) = inline_stack.last_mut() {
        builder.add_text(text);
    } else {
        current_inlines.push(Inline::Text(text));
    }
}

/// Add an inline element to the appropriate place
fn add_inline(current_inlines: &mut Vec<Inline>, inline: Inline) {
    current_inlines.push(inline);
}

/// Get parser options for pulldown-cmark
fn get_parser_options() -> Options {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);
    options
}

/// Parse code block info string
fn parse_code_block_info(info: &str) -> (Option<String>, Option<String>, Vec<u32>, bool) {
    let parts: Vec<&str> = info.split(',').collect();

    if parts.is_empty() {
        return (None, None, Vec::new(), false);
    }

    let lang = if parts[0].is_empty() {
        None
    } else {
        Some(parts[0].to_string())
    };

    let mut filename = None;
    let mut highlight_lines = Vec::new();
    let mut show_line_numbers = false;

    let mut i = 1;
    while i < parts.len() {
        let part = parts[i].trim();
        if let Some(stripped) = part.strip_prefix("filename=") {
            filename = Some(stripped.to_string());
            i += 1;
        } else if let Some(stripped) = part.strip_prefix("hl=") {
            // Collect all parts that are part of the hl= option
            let mut hl_value = stripped.to_string();
            i += 1;
            while i < parts.len() && !parts[i].contains('=') && parts[i].trim() != "ln" {
                hl_value.push(',');
                hl_value.push_str(parts[i].trim());
                i += 1;
            }
            // Parse the hl value
            for range in hl_value.split(',') {
                let range = range.trim();
                if range.contains('-') {
                    let nums: Vec<&str> = range.split('-').collect();
                    if nums.len() == 2 {
                        if let (Ok(start), Ok(end)) =
                            (nums[0].parse::<u32>(), nums[1].parse::<u32>())
                        {
                            for line in start..=end {
                                highlight_lines.push(line);
                            }
                        }
                    }
                } else if let Ok(line) = range.parse::<u32>() {
                    highlight_lines.push(line);
                }
            }
        } else if part == "ln" {
            show_line_numbers = true;
            i += 1;
        } else {
            i += 1;
        }
    }

    (lang, filename, highlight_lines, show_line_numbers)
}

/// Extract anchor ID from heading content
fn extract_anchor_id(
    content: Vec<Inline>,
    existing_id: Option<String>,
) -> (Vec<Inline>, Option<String>) {
    if existing_id.is_some() {
        return (content, existing_id);
    }

    if content.is_empty() {
        return (content, None);
    }

    if let Some(Inline::Text(text)) = content.last() {
        if let Some(anchor_start) = text.rfind("{#") {
            if let Some(anchor_end) = text[anchor_start..].find('}') {
                let anchor_id = text[anchor_start + 2..anchor_start + anchor_end].to_string();
                let mut new_content = content.clone();

                if let Inline::Text(ref mut t) = new_content
                    .last_mut()
                    .expect("last_mut should succeed after cloning")
                {
                    *t = format!(
                        "{}{}",
                        &text[..anchor_start],
                        &text[anchor_start + anchor_end + 1..]
                    );
                    *t = t.trim_end().to_string();
                }

                return (new_content, Some(anchor_id));
            }
        }
    }

    (content, None)
}

/// Extract image attributes like {width=50%} from text
fn extract_image_attributes(text: &str) -> Option<String> {
    let text = text.trim();
    if text.starts_with("{width=") && text.ends_with('}') {
        // Extract content between {width= and }
        // Length of "{width=" is 7
        if text.len() > 8 {
            let width = &text[7..text.len() - 1];
            return Some(width.to_string());
        }
    }
    None
}

/// Builder enum for constructing blocks during parsing
enum BlockBuilder {
    Heading {
        level: u8,
        content: Vec<Inline>,
        id: Option<String>,
    },
    Paragraph(Vec<Inline>),
    CodeBlock {
        lang: Option<String>,
        content: String,
        filename: Option<String>,
        highlight_lines: Vec<u32>,
        show_line_numbers: bool,
    },
    BlockQuote(Vec<Block>),
}

impl BlockBuilder {
    fn build(self) -> Block {
        match self {
            BlockBuilder::Heading { level, content, id } => Block::Heading { level, content, id },
            BlockBuilder::Paragraph(content) => Block::Paragraph(content),
            BlockBuilder::CodeBlock {
                lang,
                content,
                filename,
                highlight_lines,
                show_line_numbers,
            } => Block::CodeBlock {
                lang,
                content,
                filename,
                highlight_lines,
                show_line_numbers,
            },
            BlockBuilder::BlockQuote(content) => Block::BlockQuote(content),
        }
    }
}

/// Builder for lists
struct ListBuilder {
    ordered: bool,
    start: Option<u32>,
    items: Vec<ListItemBuilder>,
}

/// Builder for list items
struct ListItemBuilder {
    content: Vec<Block>,
    checked: Option<bool>,
}

/// Builder for tables
struct TableBuilder {
    alignments: Vec<Alignment>,
    headers: Vec<TableCell>,
    rows: Vec<Vec<TableCell>>,
    current_row: Vec<TableCell>,
    current_cell: Vec<Inline>,
}

/// Builder for inline elements
enum InlineBuilder {
    Italic(Vec<Inline>),
    Bold(Vec<Inline>),
    Strikethrough(Vec<Inline>),
    Link {
        text: Vec<Inline>,
        url: String,
        title: Option<String>,
    },
    Image {
        alt: String,
        src: String,
        title: Option<String>,
    },
}

impl InlineBuilder {
    fn add_text(&mut self, text: String) {
        match self {
            InlineBuilder::Italic(content) => content.push(Inline::Text(text)),
            InlineBuilder::Bold(content) => content.push(Inline::Text(text)),
            InlineBuilder::Strikethrough(content) => content.push(Inline::Text(text)),
            InlineBuilder::Link {
                text: link_text, ..
            } => link_text.push(Inline::Text(text)),
            InlineBuilder::Image { alt, .. } => alt.push_str(&text),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_heading() {
        let md = "# Heading 1";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Heading { level, content, id } => {
                assert_eq!(*level, 1);
                assert_eq!(content.len(), 1);
                assert_eq!(content[0], Inline::Text("Heading 1".to_string()));
                assert!(id.is_none());
            }
            _ => panic!("Expected Heading"),
        }
    }

    #[test]
    fn test_parse_heading_with_anchor() {
        let md = "# Introduction {#intro}";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Heading { level, content, id } => {
                assert_eq!(*level, 1);
                assert_eq!(content.len(), 1);
                assert_eq!(content[0], Inline::Text("Introduction".to_string()));
                assert_eq!(id, &Some("intro".to_string()));
            }
            _ => panic!("Expected Heading"),
        }
    }

    #[test]
    fn test_parse_paragraph() {
        let md = "This is a paragraph.";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert_eq!(content.len(), 1);
                assert_eq!(content[0], Inline::Text("This is a paragraph.".to_string()));
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_bold_italic() {
        let md = "This is **bold** and *italic* and ***both***.";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert!(!content.is_empty());
                assert!(content.iter().any(|i| matches!(i, Inline::Bold(_))));
                assert!(content.iter().any(|i| matches!(i, Inline::Italic(_))));
                // Note: ***both*** might be parsed as Bold(Italic(...)) or Italic(Bold(...))
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_inline_code() {
        let md = "Use `println!` for output.";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert!(content.iter().any(|i| matches!(i, Inline::Code(_))));
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_code_block() {
        let md = "```rust\nfn main() {\n    println!(\"Hello\");\n}\n```";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::CodeBlock {
                lang,
                content,
                filename,
                highlight_lines,
                show_line_numbers,
            } => {
                assert_eq!(lang, &Some("rust".to_string()));
                assert!(content.contains("println!"));
                assert!(filename.is_none());
                assert!(highlight_lines.is_empty());
                assert!(!show_line_numbers);
            }
            _ => panic!("Expected CodeBlock"),
        }
    }

    #[test]
    fn test_parse_code_block_with_options() {
        let md = "```rust,filename=main.rs,hl=2,4-5,ln\nfn main() {\n    println!(\"Hello\");\n    let x = 5;\n    let y = 10;\n}\n```";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::CodeBlock {
                lang,
                content: _,
                filename,
                highlight_lines,
                show_line_numbers,
            } => {
                assert_eq!(lang, &Some("rust".to_string()));
                assert_eq!(filename, &Some("main.rs".to_string()));
                assert_eq!(highlight_lines, &vec![2, 4, 5]);
                assert!(*show_line_numbers);
            }
            _ => panic!("Expected CodeBlock"),
        }
    }

    #[test]
    fn test_parse_mermaid_block() {
        let md = "```mermaid\nflowchart LR\n    A --> B\n```";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Mermaid { content, id } => {
                assert!(content.contains("flowchart"));
                assert!(id.is_none());
            }
            _ => panic!("Expected Mermaid"),
        }
    }

    #[test]
    fn test_parse_blockquote() {
        let md = "> This is a quote\n> with multiple lines";
        let doc = parse_markdown(md);
        assert!(!doc.blocks.is_empty());
        assert!(doc.blocks.iter().any(|b| matches!(b, Block::BlockQuote(_))));
    }

    #[test]
    fn test_parse_unordered_list() {
        let md = "- Item 1\n- Item 2\n- Item 3";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::List {
                ordered,
                start,
                items,
            } => {
                assert!(!ordered);
                assert!(start.is_none());
                assert_eq!(items.len(), 3);
            }
            _ => panic!("Expected List"),
        }
    }

    #[test]
    fn test_parse_ordered_list() {
        let md = "1. First\n2. Second\n3. Third";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::List { ordered, items, .. } => {
                assert!(ordered);
                assert_eq!(items.len(), 3);
            }
            _ => panic!("Expected List"),
        }
    }

    #[test]
    fn test_parse_task_list() {
        let md = "- [x] Done\n- [ ] Not done\n- [x] Also done";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::List { ordered, items, .. } => {
                assert!(!ordered);
                assert_eq!(items.len(), 3);
                assert_eq!(items[0].checked, Some(true));
                assert_eq!(items[1].checked, Some(false));
                assert_eq!(items[2].checked, Some(true));
            }
            _ => panic!("Expected List"),
        }
    }

    #[test]
    fn test_parse_table_with_comment_id() {
        let md = "<!-- {#tbl:users} -->\n| Name | Email |\n|------|-------|\n| John | john@example.com |";
        let doc = parse_markdown(md);
        // doc.blocks[0] should be the Table
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Table { id, .. } => {
                assert_eq!(id.as_deref(), Some("tbl:users"));
            }
            _ => panic!("Expected Table, found {:?}", doc.blocks),
        }
    }

    #[test]
    fn test_parse_table_with_caption_id() {
        let md = "Table: User List {#tbl:users}\n| Name | Email |\n|------|-------|\n| John | john@example.com |";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Table { caption, id, .. } => {
                assert_eq!(caption.as_deref(), Some("User List"));
                assert_eq!(id.as_deref(), Some("tbl:users"));
            }
            _ => panic!("Expected Table"),
        }
    }

    #[test]
    fn test_parse_table_with_caption_no_id() {
        let md = "Table: My Caption\n| Col 1 |\n|-------|\n| val |";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Table { caption, id, .. } => {
                assert_eq!(caption.as_deref(), Some("My Caption"));
                assert!(id.is_none());
            }
            _ => panic!("Expected Table"),
        }
    }

    #[test]
    fn test_parse_thematic_break() {
        let md = "---";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        assert!(matches!(doc.blocks[0], Block::ThematicBreak));
    }

    #[test]
    fn test_parse_link() {
        let md = "[OpenAI](https://openai.com)";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert_eq!(content.len(), 1);
                assert!(matches!(content[0], Inline::Link { .. }));
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_strikethrough() {
        let md = "~~deleted text~~";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert_eq!(content.len(), 1);
                assert!(matches!(content[0], Inline::Strikethrough(_)));
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_footnote_reference() {
        let md = "Text with footnote[^1]\n\n[^1]: This is the footnote";
        let doc = parse_markdown(md);
        assert!(!doc.blocks.is_empty());
        // Check if any paragraph contains a footnote reference
        let has_footnote = doc.blocks.iter().any(|b| {
            if let Block::Paragraph(content) = b {
                content.iter().any(|i| matches!(i, Inline::FootnoteRef(_)))
            } else {
                false
            }
        });
        assert!(has_footnote, "Expected to find a footnote reference");
    }

    #[test]
    fn test_parse_footnote_definition() {
        let md = "Text with footnote[^1]\n\n[^1]: This is the footnote content.";
        let doc = parse_markdown(md);
        // Check that footnotes map contains the definition
        assert!(
            doc.footnotes.contains_key("1"),
            "Expected footnote '1' to be defined"
        );
        let footnote_content = &doc.footnotes["1"];
        assert!(
            !footnote_content.is_empty(),
            "Expected footnote content to not be empty"
        );
        // Check that the footnote contains a paragraph with the text
        assert!(footnote_content.iter().any(|b| {
            if let Block::Paragraph(content) = b {
                content.iter().any(|i| matches!(i, Inline::Text(_)))
            } else {
                false
            }
        }));
    }

    #[test]
    fn test_parse_multiple_footnotes() {
        let md = "Text with footnotes[^1][^2]\n\n[^1]: First footnote\n[^2]: Second footnote";
        let doc = parse_markdown(md);
        assert_eq!(doc.footnotes.len(), 2, "Expected 2 footnotes");
        assert!(doc.footnotes.contains_key("1"));
        assert!(doc.footnotes.contains_key("2"));
    }

    #[test]
    fn test_parse_footnote_with_multiple_paragraphs() {
        let md = "Text with footnote[^1]\n\n[^1]: First paragraph.\n\nSecond paragraph.";
        let doc = parse_markdown(md);
        assert!(doc.footnotes.contains_key("1"));
        let footnote_content = &doc.footnotes["1"];
        assert_eq!(
            footnote_content.len(),
            2,
            "Expected 2 paragraphs in footnote"
        );
    }

    #[test]
    fn test_parse_multiple_blocks() {
        let md = "# Title\n\nParagraph 1\n\nParagraph 2";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 3);
        assert!(matches!(doc.blocks[0], Block::Heading { .. }));
        assert!(matches!(doc.blocks[1], Block::Paragraph(_)));
        assert!(matches!(doc.blocks[2], Block::Paragraph(_)));
    }

    #[test]
    fn test_parse_code_block_info() {
        let (lang, filename, highlight_lines, show_line_numbers) = parse_code_block_info("rust");
        assert_eq!(lang, Some("rust".to_string()));
        assert!(filename.is_none());
        assert!(highlight_lines.is_empty());
        assert!(!show_line_numbers);

        let (lang, filename, highlight_lines, show_line_numbers) =
            parse_code_block_info("rust,filename=main.rs,hl=3,5-7,ln");
        assert_eq!(lang, Some("rust".to_string()));
        assert_eq!(filename, Some("main.rs".to_string()));
        assert_eq!(highlight_lines, vec![3, 5, 6, 7]);
        assert!(show_line_numbers);
    }

    #[test]
    fn test_extract_anchor_id() {
        let content = vec![Inline::Text("Introduction {#intro}".to_string())];
        let (new_content, id) = extract_anchor_id(content, None);
        assert_eq!(id, Some("intro".to_string()));
        assert_eq!(new_content.len(), 1);
        assert_eq!(new_content[0], Inline::Text("Introduction".to_string()));
    }

    #[test]
    fn test_extract_anchor_id_no_anchor() {
        let content = vec![Inline::Text("Introduction".to_string())];
        let (new_content, id) = extract_anchor_id(content, None);
        assert!(id.is_none());
        assert_eq!(new_content.len(), 1);
        assert_eq!(new_content[0], Inline::Text("Introduction".to_string()));
    }

    #[test]
    fn test_complex_document() {
        let md = r#"# Getting Started {#ch01}

Welcome to the documentation!

## Prerequisites

You need:
- Rust 1.75+
- 4GB RAM

## Example Code

```rust,filename=main.rs,hl=2
fn main() {
    println!("Hello, world!");
}
```

## Features

| Feature | Status |
|---------|--------|
| Tables  | ✅      |
| Code    | ✅      |

> **Note**: This is important.

~~Deprecated feature~~ has been removed.
"#;

        let doc = parse_markdown(md);
        assert!(!doc.blocks.is_empty());

        // Check for heading with anchor
        assert!(doc
            .blocks
            .iter()
            .any(|b| matches!(b, Block::Heading { id: Some(_), .. })));

        // Check for code block with options
        assert!(doc.blocks.iter().any(|b| matches!(
            b,
            Block::CodeBlock {
                filename: Some(_),
                ..
            }
        )));

        // Check for table
        assert!(doc.blocks.iter().any(|b| matches!(b, Block::Table { .. })));

        // Check for blockquote
        assert!(doc.blocks.iter().any(|b| matches!(b, Block::BlockQuote(_))));
    }

    #[test]
    fn test_parse_image_with_width() {
        let md = "![Image](image.png){width=50%}";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Image { width, .. } => {
                assert_eq!(width, &Some("50%".to_string()));
            }
            _ => panic!("Expected Image block with width"),
        }
    }

    #[test]
    fn test_parse_image_with_width_and_space() {
        let md = "![Image](image.png) {width=800px}";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Image { width, .. } => {
                assert_eq!(width, &Some("800px".to_string()));
            }
            _ => panic!("Expected Image block with width"),
        }
    }

    #[test]
    fn test_parse_image_simple() {
        let md = "![Image](image.png)";
        let doc = parse_markdown(md);
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Image { width, .. } => {
                assert!(width.is_none());
            }
            _ => panic!("Expected Image block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_parse_image_surrounded() {
        let md = "\n\n![Image](image.png)\n\n";
        let doc = parse_markdown(md);
        assert!(!doc.blocks.is_empty());
        match &doc.blocks[0] {
            Block::Image { width, .. } => {
                assert!(width.is_none());
            }
            _ => panic!("Expected Image block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_parse_cross_reference_simple() {
        let md = "See {ref:intro} for details.";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                assert!(content
                    .iter()
                    .any(|i| matches!(i, Inline::CrossRef { target, .. } if target == "intro")));
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_cross_reference_with_type() {
        let md = "See {ref:fig:diagram} for the architecture.";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                let has_fig_ref = content.iter().any(|i| {
                    matches!(i, Inline::CrossRef { target, ref_type }
                        if target == "diagram" && *ref_type == RefType::Figure)
                });
                assert!(has_fig_ref, "Expected figure cross-reference");
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_multiple_cross_references() {
        let md = "See {ref:ch01} and {ref:fig:arch}.";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                let cross_ref_count = content
                    .iter()
                    .filter(|i| matches!(i, Inline::CrossRef { .. }))
                    .count();
                assert_eq!(cross_ref_count, 2);
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_cross_reference_in_bold() {
        let md = "**See {ref:intro} for more**";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::Paragraph(content) => {
                if let Inline::Bold(inner) = &content[0] {
                    assert!(inner.iter().any(|i| matches!(i, Inline::CrossRef { .. })));
                } else {
                    panic!("Expected Bold inline");
                }
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn test_parse_include_directive() {
        let md = "{!include:chapters/intro.md}";
        let doc = parse_markdown(md);

        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Include { path, .. } => {
                assert_eq!(path, "chapters/intro.md");
            }
            _ => panic!("Expected Include block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_parse_code_include_directive() {
        let md = "{!code:src/main.rs}";
        let doc = parse_markdown(md);

        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::CodeInclude {
                path,
                start_line,
                end_line,
                lang,
            } => {
                assert_eq!(path, "src/main.rs");
                assert!(start_line.is_none());
                assert!(end_line.is_none());
                assert!(lang.is_none());
            }
            _ => panic!("Expected CodeInclude block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_parse_code_include_with_lines() {
        let md = "{!code:src/main.rs:10-25}";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::CodeInclude {
                path,
                start_line,
                end_line,
                lang,
            } => {
                assert_eq!(path, "src/main.rs");
                assert_eq!(*start_line, Some(10));
                assert_eq!(*end_line, Some(25));
                assert!(lang.is_none());
            }
            _ => panic!("Expected CodeInclude block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_parse_code_include_with_lang() {
        let md = "{!code:src/config.txt:5-15:yaml}";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::CodeInclude {
                path,
                start_line,
                end_line,
                lang,
            } => {
                assert_eq!(path, "src/config.txt");
                assert_eq!(*start_line, Some(5));
                assert_eq!(*end_line, Some(15));
                assert_eq!(lang, &Some("yaml".to_string()));
            }
            _ => panic!("Expected CodeInclude block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_include_not_matched_in_text() {
        // Include directive in the middle of text should NOT be converted
        let md = "See {!include:file.md} for more info.";
        let doc = parse_markdown(md);

        // Should remain as a paragraph, not be converted to Include
        assert_eq!(doc.blocks.len(), 1);
        assert!(matches!(doc.blocks[0], Block::Paragraph(_)));
    }

    #[test]
    fn test_include_directive_with_surrounding_whitespace() {
        let md = "\n\n{!include:chapters/intro.md}\n\n";
        let doc = parse_markdown(md);

        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Include { path, .. } => {
                assert_eq!(path, "chapters/intro.md");
            }
            _ => panic!("Expected Include block, found {:?}", doc.blocks[0]),
        }
    }

    #[test]
    fn test_code_include_without_lines() {
        let md = "{!code:src/main.rs:rust}";
        let doc = parse_markdown(md);

        match &doc.blocks[0] {
            Block::CodeInclude {
                path,
                start_line,
                end_line,
                lang,
            } => {
                assert_eq!(path, "src/main.rs");
                assert!(start_line.is_none());
                assert!(end_line.is_none());
                assert_eq!(lang, &Some("rust".to_string()));
            }
            _ => panic!("Expected CodeInclude block, found {:?}", doc.blocks[0]),
        }
    }
}
