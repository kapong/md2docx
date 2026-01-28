//! Table template rendering
//!
//! Applies extracted table styles to markdown tables during document generation.

use crate::docx::ooxml::{Paragraph, Run, Table, TableCellElement, TableRow, TableWidth};
use crate::error::Result;
use crate::template::extract::table::TableTemplate;

/// Render a table with extracted template styles
///
/// This function takes a markdown table and applies the styles from
/// the extracted table template.
///
/// # Arguments
/// * `headers` - Table header cells (content strings)
/// * `rows` - Table data rows (content strings)
/// * `template` - The extracted table template
/// * `caption` - Optional caption text
///
/// # Returns
/// A `Table` element with applied styles
///
/// # Example
/// ```rust,no_run
/// use md2docx::template::extract::TableTemplate;
/// use md2docx::template::render::table::render_table_with_template;
///
/// let template = TableTemplate::default();
/// let headers = vec!["Name".to_string(), "Type".to_string()];
/// let rows = vec![
///     vec!["Admin".to_string(), "String".to_string()],
///     vec!["User".to_string(), "String".to_string()],
/// ];
///
/// let table = render_table_with_template(&headers, &rows, &template, Some("User Types")).unwrap();
/// ```
pub fn render_table_with_template(
    headers: &[String],
    rows: &[Vec<String>],
    template: &TableTemplate,
    _caption: Option<&str>,
) -> Result<Table> {
    let mut table = Table::new().with_header_row(true);

    // Calculate column count from headers
    let col_count = headers.len();
    if col_count == 0 {
        return Ok(table);
    }

    // Set table width to 100%
    table = table.width(TableWidth::Pct(5000));

    // Set column widths (equal distribution)
    let col_width = 9000 / col_count as u32;
    table = table.with_column_widths(vec![col_width; col_count]);

    // Add caption paragraph if provided
    // Note: In actual implementation, caption would be added before the table
    // in the document, not inside the table itself

    // Add header row
    let header_style = &template.header;
    let mut header_row = TableRow::new().header();

    for (col_idx, header_content) in headers.iter().enumerate() {
        let cell_style = if col_idx == 0 {
            &template.first_column
        } else {
            &template.other_columns
        };

        let para = create_styled_paragraph(
            header_content.as_str(),
            cell_style.font_family.clone(),
            cell_style.font_size,
            cell_style.font_color.clone(),
            header_style.bold || cell_style.bold,
            header_style.italic || cell_style.italic,
        );

        let cell = TableCellElement::new()
            .width(TableWidth::Pct(5000 / col_count as u32))
            .add_paragraph(para);

        // Note: Background color would be set here in actual implementation
        header_row = header_row.add_cell(cell);
    }

    table = table.add_row(header_row);

    // Add data rows
    for (row_idx, row_data) in rows.iter().enumerate() {
        let row_style = template.row_style_for_index(row_idx + 1); // +1 because header is row 0
        let mut data_row = TableRow::new();

        for (col_idx, cell_content) in row_data.iter().enumerate() {
            let cell_style = template.cell_style_for_column(col_idx);

            let para = create_styled_paragraph(
                cell_content.as_str(),
                cell_style.font_family.clone(),
                cell_style.font_size,
                cell_style.font_color.clone(),
                row_style.bold || cell_style.bold,
                row_style.italic || cell_style.italic,
            );

            let cell = TableCellElement::new()
                .width(TableWidth::Pct(5000 / col_count as u32))
                .add_paragraph(para);

            // Note: Background color would be set here in actual implementation
            data_row = data_row.add_cell(cell);
        }

        table = table.add_row(data_row);
    }

    Ok(table)
}

/// Create a styled paragraph for table cells
fn create_styled_paragraph(
    text: &str,
    font_family: String,
    font_size: u32,
    font_color: String,
    bold: bool,
    italic: bool,
) -> Paragraph {
    let mut run = Run::new(text);

    // Apply font
    run.font = Some(font_family);

    // Apply size
    run.size = Some(font_size);

    // Apply color
    run.color = Some(font_color);

    // Apply bold
    run.bold = bold;

    // Apply italic
    run.italic = italic;

    Paragraph::new().add_run(run)
}

/// Format a table caption with the template style
///
/// # Arguments
/// * `template` - The table template with caption style
/// * `number` - The table number (e.g., "1.2")
/// * `text` - The caption text
///
/// # Returns
/// A formatted caption string
pub fn format_table_caption(template: &TableTemplate, number: &str, text: &str) -> String {
    format!("{} {}: {}", template.caption.prefix, number, text)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_table_caption() {
        let template = TableTemplate::default();
        let caption = format_table_caption(&template, "1.2", "User account types");

        assert_eq!(caption, "Table 1.2: User account types");
    }

    #[test]
    fn test_render_table_with_template() {
        let template = TableTemplate::default();
        let headers = vec!["Name".to_string(), "Type".to_string()];
        let rows = vec![
            vec!["Admin".to_string(), "String".to_string()],
            vec!["User".to_string(), "String".to_string()],
        ];

        let table =
            render_table_with_template(&headers, &rows, &template, Some("User Types")).unwrap();

        assert_eq!(table.rows.len(), 3); // Header + 2 data rows
        assert!(table.has_header_row);
    }

    #[test]
    fn test_render_empty_table() {
        let template = TableTemplate::default();
        let headers: Vec<String> = vec![];
        let rows: Vec<Vec<String>> = vec![];

        let table = render_table_with_template(&headers, &rows, &template, None).unwrap();

        assert!(table.rows.is_empty());
    }
}
