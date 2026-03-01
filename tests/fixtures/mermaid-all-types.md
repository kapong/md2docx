# Mermaid Diagram Type Test

This document tests all common Mermaid diagram types to verify which ones render as PNG images and which fall back to code blocks.

---

## 1. Flowchart (TD) — Supported

```mermaid
flowchart TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Action 1]
    B -->|No| D[Action 2]
    C --> E[End]
    D --> E
```

## 2. Flowchart (LR) — Supported

```mermaid
flowchart LR
    Input([Input]) --> Validate[[Validate]]
    Validate --> Store[(Database)]
    Store --> Output>Result]
    Output --> Done{{Done}}
```

## 3. Graph (TD) — Supported

```mermaid
graph TD
    A[Module A] --> B[Module B]
    A --> C[Module C]
    B --> D[Module D]
    C --> D
    D --> E[Output]
```

## 4. Graph (LR) — Supported

```mermaid
graph LR
    Start --> Stop
    Start --> Process
    Process --> Stop
```

## 5. Flowchart with Subgraphs — Supported

```mermaid
flowchart TB
    subgraph Frontend
        A[React App] --> B[API Client]
    end
    subgraph Backend
        C[REST API] --> D[Service Layer]
        D --> E[(PostgreSQL)]
    end
    B --> C
```

## 6. Flowchart with Styles — Supported

```mermaid
flowchart LR
    A[Normal] --> B[Important]
    B --> C[Critical]
    style B fill:#f9f,stroke:#333,stroke-width:2px
    style C fill:#f66,stroke:#900,stroke-width:4px
```

---

## 7. Sequence Diagram — Unsupported (fallback to code)

```mermaid
sequenceDiagram
    participant Client
    participant Server
    participant Database

    Client->>Server: HTTP Request
    Server->>Database: SQL Query
    Database-->>Server: Result Set
    Server-->>Client: JSON Response

    Note over Client,Server: REST API Communication
```

## 8. ER Diagram — Unsupported (fallback to code)

```mermaid
erDiagram
    CUSTOMER ||--o{ ORDER : places
    ORDER ||--|{ LINE-ITEM : contains
    PRODUCT ||--o{ LINE-ITEM : "ordered in"
    CUSTOMER {
        string name
        string email
        int id PK
    }
    ORDER {
        int id PK
        date created
        string status
    }
    PRODUCT {
        int id PK
        string name
        float price
    }
```

## 9. Class Diagram — Unsupported (fallback to code)

```mermaid
classDiagram
    class Document {
        +String title
        +String content
        +render() Vec~u8~
    }
    class Parser {
        +parse(md: String) Document
    }
    class Builder {
        +build(doc: Document) Docx
    }
    Parser --> Document : creates
    Builder --> Document : consumes
    Document <|-- MarkdownDoc
    Document <|-- HtmlDoc
```

## 10. State Diagram — Unsupported (fallback to code)

```mermaid
stateDiagram-v2
    [*] --> Draft
    Draft --> Review : Submit
    Review --> Approved : Approve
    Review --> Draft : Reject
    Approved --> Published : Publish
    Published --> [*]

    state Review {
        [*] --> Pending
        Pending --> InReview
        InReview --> [*]
    }
```

## 11. Gantt Chart — Unsupported (fallback to code)

```mermaid
gantt
    title Development Schedule
    dateFormat YYYY-MM-DD
    axisFormat %b %d

    section Design
    Requirements     :done, des1, 2026-01-01, 14d
    Architecture     :done, des2, after des1, 7d

    section Implementation
    Core Engine      :active, impl1, after des2, 21d
    PNG Support      :impl2, after impl1, 7d
    Testing          :impl3, after impl2, 14d

    section Release
    Documentation    :doc1, after impl2, 10d
    Release v1.0     :milestone, after impl3, 0d
```

## 12. Pie Chart — Unsupported (fallback to code)

```mermaid
pie title Mermaid Diagram Support
    "Flowchart/Graph" : 40
    "Unsupported (code fallback)" : 60
```

## 13. User Journey — Unsupported (fallback to code)

```mermaid
journey
    title User Document Conversion Journey
    section Writing
      Write Markdown: 5: Author
      Add Diagrams: 4: Author
    section Converting
      Run md2docx: 5: Author, Tool
      Review Output: 3: Author
    section Sharing
      Send DOCX: 5: Author
      Read Document: 4: Reader
```

## 14. Git Graph — Unsupported (fallback to code)

```mermaid
gitGraph
    commit
    commit
    branch develop
    checkout develop
    commit
    commit
    checkout main
    merge develop
    commit
    branch feature
    checkout feature
    commit
    checkout develop
    merge feature
    checkout main
    merge develop
```

## 15. Mindmap — Unsupported (fallback to code)

```mermaid
mindmap
    root((md2docx))
        Input
            Markdown
            TOML Config
            Template DOCX
        Processing
            Parser
            Builder
            Mermaid Renderer
        Output
            DOCX File
            Embedded Fonts
            PNG Images
```

## 16. Timeline — Unsupported (fallback to code)

```mermaid
timeline
    title md2docx Version History
    2025-01 : v0.1.0 : Initial release
    2025-06 : v0.2.0 : Mermaid SVG support
    2026-01 : v0.2.2 : Math equation images
    2026-03 : v0.3.0 : Mermaid PNG output
```

## 17. Quadrant Chart — Unsupported (fallback to code)

```mermaid
quadrantChart
    title Feature Priority Matrix
    x-axis Low Effort --> High Effort
    y-axis Low Impact --> High Impact
    quadrant-1 Do First
    quadrant-2 Plan
    quadrant-3 Delegate
    quadrant-4 Eliminate
    PNG Output: [0.2, 0.9]
    SVG Fallback: [0.1, 0.6]
    Sequence Diagrams: [0.8, 0.7]
    Theme Support: [0.6, 0.4]
```

## 18. Sankey Diagram — Unsupported (fallback to code)

```mermaid
sankey-beta
    Markdown,Parser,100
    Parser,Builder,80
    Parser,Metadata,20
    Builder,DOCX,70
    Builder,Images,10
```

## 19. Block Diagram — Unsupported (fallback to code)

```mermaid
block-beta
    columns 3
    Frontend:3
    A["API"] B["Auth"] C["Cache"]
    Backend:3
    D["Service"] E["Queue"] F["DB"]
```

---

## Summary

| # | Diagram Type | Status |
|---|-------------|--------|
| 1 | flowchart TD | **Rendered as PNG** |
| 2 | flowchart LR | **Rendered as PNG** |
| 3 | graph TD | **Rendered as PNG** |
| 4 | graph LR | **Rendered as PNG** |
| 5 | flowchart (subgraphs) | **Rendered as PNG** |
| 6 | flowchart (styles) | **Rendered as PNG** |
| 7 | sequenceDiagram | Fallback to code |
| 8 | erDiagram | Fallback to code |
| 9 | classDiagram | Fallback to code |
| 10 | stateDiagram | Fallback to code |
| 11 | gantt | Fallback to code |
| 12 | pie | Fallback to code |
| 13 | journey | Fallback to code |
| 14 | gitGraph | Fallback to code |
| 15 | mindmap | Fallback to code |
| 16 | timeline | Fallback to code |
| 17 | quadrantChart | Fallback to code |
| 18 | sankey | Fallback to code |
| 19 | block-beta | Fallback to code |
