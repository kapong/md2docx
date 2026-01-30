# Mermaid Diagram Demo

This document demonstrates various Mermaid diagram types rendered by md2docx.

## Flowchart

```mermaid
flowchart TD
    A[Start] --> B(Process)
    B --> C{Is it working?}
    C --> D[Great!]
    C --> E[Debug]
    D --> F[Deploy]
    E --> B
```

## Sequence Diagram

```mermaid
sequenceDiagram
    participant User
    participant App
    participant Database
    
    User->>App: Login request
    App->>Database: Validate credentials
    Database-->>App: User data
    App-->>User: Login success
```

## Class Diagram

```mermaid
classDiagram
    class Animal {
        +String name
        +makeSound()
    }
    class Dog {
        +fetch()
    }
    class Cat {
        +climb()
    }
    Animal <|-- Dog
    Animal <|-- Cat
```

## State Diagram

```mermaid
stateDiagram-v2
    [*] --> Idle
    Idle --> Processing
    Processing --> Success
    Processing --> Error
    Success --> [*]
    Error --> Idle
```

## Gantt Chart

```mermaid
gantt
    title Project Timeline
    dateFormat  YYYY-MM-DD
    section Planning
    Requirements    :done, a1, 2024-01-01, 7d
    Design          :active, a2, after a1, 5d
    section Development
    Implementation  :a3, after a2, 14d
    Testing         :a4, after a3, 7d
```

## Flowchart Shapes

```mermaid
flowchart LR
    A([Start]) --> B[[Process]]
    B --> C[(Database)]
    C --> D>Output]
    D --> E{{Done}}
```
