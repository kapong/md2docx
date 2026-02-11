---
title: "Code Syntax Highlighting Demo"
language: en
---

# Code Syntax Highlighting

md2docx supports syntax highlighting for code blocks. Specify the language after the opening triple backticks to enable highlighting.

## Rust

```rust
use std::collections::HashMap;

#[derive(Debug, Clone)]
struct Config {
    name: String,
    values: HashMap<String, i32>,
}

impl Config {
    fn new(name: &str) -> Self {
        Self {
            name: name.to_string(),
            values: HashMap::new(),
        }
    }

    fn get(&self, key: &str) -> Option<&i32> {
        self.values.get(key)
    }
}

fn main() {
    let mut config = Config::new("app");
    config.values.insert("port".to_string(), 8080);

    if let Some(port) = config.get("port") {
        println!("Server running on port {}", port);
    }
}
```

## Python

```python
from dataclasses import dataclass
from typing import Optional
import asyncio

@dataclass
class User:
    name: str
    email: str
    age: int = 0

    @property
    def is_adult(self) -> bool:
        return self.age >= 18

async def fetch_users(url: str) -> list[User]:
    """Fetch users from the API."""
    users = [
        User("Alice", "alice@example.com", 30),
        User("Bob", "bob@example.com", 17),
    ]
    await asyncio.sleep(0.1)  # simulate network delay
    return [u for u in users if u.is_adult]

if __name__ == "__main__":
    result = asyncio.run(fetch_users("https://api.example.com"))
    for user in result:
        print(f"{user.name}: {user.email}")
```

## JavaScript

```javascript
class EventEmitter {
  #listeners = new Map();

  on(event, callback) {
    if (!this.#listeners.has(event)) {
      this.#listeners.set(event, []);
    }
    this.#listeners.get(event).push(callback);
    return this;
  }

  emit(event, ...args) {
    const callbacks = this.#listeners.get(event) ?? [];
    callbacks.forEach((cb) => cb(...args));
  }
}

// Usage
const emitter = new EventEmitter();
emitter.on("data", (msg) => console.log(`Received: ${msg}`));
emitter.emit("data", "Hello, world!");
```

## TypeScript

```typescript
interface Repository<T> {
  findById(id: string): Promise<T | null>;
  save(entity: T): Promise<T>;
  delete(id: string): Promise<boolean>;
}

type User = {
  id: string;
  name: string;
  email: string;
  createdAt: Date;
};

class UserRepository implements Repository<User> {
  private users: Map<string, User> = new Map();

  async findById(id: string): Promise<User | null> {
    return this.users.get(id) ?? null;
  }

  async save(user: User): Promise<User> {
    this.users.set(user.id, user);
    return user;
  }

  async delete(id: string): Promise<boolean> {
    return this.users.delete(id);
  }
}
```

## Go

```go
package main

import (
	"fmt"
	"net/http"
	"sync"
)

type Counter struct {
	mu    sync.Mutex
	count int
}

func (c *Counter) Increment() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.count++
}

func (c *Counter) Value() int {
	c.mu.Lock()
	defer c.mu.Unlock()
	return c.count
}

func main() {
	counter := &Counter{}

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		counter.Increment()
		fmt.Fprintf(w, "Visits: %d\n", counter.Value())
	})

	fmt.Println("Server starting on :8080")
	http.ListenAndServe(":8080", nil)
}
```

## Java

```java
import java.util.*;
import java.util.stream.*;

public record Person(String name, int age, String city) {

    public static List<Person> filterAdults(List<Person> people) {
        return people.stream()
            .filter(p -> p.age() >= 18)
            .sorted(Comparator.comparing(Person::name))
            .toList();
    }

    public static Map<String, List<Person>> groupByCity(List<Person> people) {
        return people.stream()
            .collect(Collectors.groupingBy(Person::city));
    }

    public static void main(String[] args) {
        var people = List.of(
            new Person("Alice", 30, "Tokyo"),
            new Person("Bob", 17, "London"),
            new Person("Charlie", 25, "Tokyo")
        );

        var adults = filterAdults(people);
        adults.forEach(p -> System.out.printf("%s (%d)%n", p.name(), p.age()));
    }
}
```

## C

```c
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef struct Node {
    int data;
    struct Node *next;
} Node;

Node *create_node(int data) {
    Node *node = malloc(sizeof(Node));
    if (!node) return NULL;
    node->data = data;
    node->next = NULL;
    return node;
}

void push(Node **head, int data) {
    Node *node = create_node(data);
    node->next = *head;
    *head = node;
}

void print_list(const Node *head) {
    while (head) {
        printf("%d -> ", head->data);
        head = head->next;
    }
    printf("NULL\n");
}

int main(void) {
    Node *list = NULL;
    for (int i = 5; i >= 1; i--) {
        push(&list, i);
    }
    print_list(list);  // 1 -> 2 -> 3 -> 4 -> 5 -> NULL
    return 0;
}
```

## C++

```cpp
#include <iostream>
#include <vector>
#include <algorithm>
#include <ranges>

template <typename T>
concept Printable = requires(T t) {
    { std::cout << t } -> std::same_as<std::ostream&>;
};

template <Printable T>
void print_sorted(std::vector<T> items) {
    std::ranges::sort(items);
    for (const auto& item : items) {
        std::cout << item << " ";
    }
    std::cout << "\n";
}

class Matrix {
    std::vector<std::vector<double>> data_;
public:
    Matrix(int rows, int cols)
        : data_(rows, std::vector<double>(cols, 0.0)) {}

    double& operator()(int r, int c) { return data_[r][c]; }
    const double& operator()(int r, int c) const { return data_[r][c]; }
};

int main() {
    std::vector<int> nums = {5, 3, 1, 4, 2};
    print_sorted(nums);  // 1 2 3 4 5

    Matrix m(3, 3);
    m(1, 1) = 42.0;
    std::cout << "m[1][1] = " << m(1, 1) << std::endl;
}
```

## C#

```cs
using System;
using System.Linq;
using System.Collections.Generic;

namespace Demo;

public record Product(string Name, decimal Price, string Category);

public static class Store
{
    public static IEnumerable<IGrouping<string, Product>> GroupByCategory(
        IEnumerable<Product> products) =>
        products
            .Where(p => p.Price > 0)
            .OrderBy(p => p.Name)
            .GroupBy(p => p.Category);

    public static void Main()
    {
        var products = new List<Product>
        {
            new("Laptop", 999.99m, "Electronics"),
            new("Keyboard", 79.99m, "Electronics"),
            new("Coffee", 12.99m, "Food"),
        };

        foreach (var group in GroupByCategory(products))
        {
            Console.WriteLine($"== {group.Key} ==");
            foreach (var p in group)
                Console.WriteLine($"  {p.Name}: ${p.Price}");
        }
    }
}
```

## Ruby

```ruby
module Greeting
  def greet(name)
    "Hello, #{name}!"
  end
end

class Person
  include Greeting
  attr_accessor :name, :age

  def initialize(name, age)
    @name = name
    @age = age
  end

  def adult?
    age >= 18
  end

  def <=>(other)
    name <=> other.name
  end
end

people = [
  Person.new("Charlie", 25),
  Person.new("Alice", 30),
  Person.new("Bob", 17),
]

adults = people.select(&:adult?).sort_by(&:name)
adults.each { |p| puts "#{p.greet(p.name)} (age #{p.age})" }
```

## Shell / Bash

```bash
#!/usr/bin/env bash
set -euo pipefail

readonly LOG_FILE="/var/log/deploy.log"

log() {
    local level="$1"; shift
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] [$level] $*" | tee -a "$LOG_FILE"
}

deploy() {
    local env="${1:?Usage: deploy <environment>}"
    log INFO "Starting deployment to $env"

    if [[ "$env" == "production" ]]; then
        read -rp "Are you sure? (y/N) " confirm
        [[ "$confirm" =~ ^[Yy]$ ]] || { log WARN "Aborted"; exit 1; }
    fi

    for service in api worker web; do
        log INFO "Deploying $service..."
        # docker compose -f "compose.$env.yaml" up -d "$service"
        sleep 1
    done

    log INFO "Deployment to $env complete"
}

deploy "$@"
```

## SQL

```sql
CREATE TABLE employees (
    id         SERIAL PRIMARY KEY,
    name       VARCHAR(100) NOT NULL,
    department VARCHAR(50),
    salary     DECIMAL(10, 2),
    hired_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Find average salary by department for departments with more than 5 employees
SELECT
    department,
    COUNT(*)          AS employee_count,
    AVG(salary)       AS avg_salary,
    MAX(salary)       AS max_salary
FROM employees
WHERE salary > 0
GROUP BY department
HAVING COUNT(*) > 5
ORDER BY avg_salary DESC;

-- Recursive CTE: org chart
WITH RECURSIVE org_chart AS (
    SELECT id, name, manager_id, 1 AS depth
    FROM employees WHERE manager_id IS NULL
    UNION ALL
    SELECT e.id, e.name, e.manager_id, oc.depth + 1
    FROM employees e
    JOIN org_chart oc ON e.manager_id = oc.id
)
SELECT * FROM org_chart ORDER BY depth, name;
```

## HTML

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard</title>
    <style>
        .card {
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 1rem;
            margin: 0.5rem;
        }
        .card h3 { color: #333; }
    </style>
</head>
<body>
    <main id="app">
        <h1>Welcome</h1>
        <div class="card">
            <h3>Status</h3>
            <p>All systems operational.</p>
        </div>
    </main>
    <script>
        document.querySelector('#app h1').textContent = 'Dashboard';
    </script>
</body>
</html>
```

## CSS

```css
:root {
    --primary: #3b82f6;
    --radius: 0.5rem;
    --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
}

.container {
    max-width: 1200px;
    margin: 0 auto;
    padding: 0 1rem;
}

.btn {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.5rem 1rem;
    border: none;
    border-radius: var(--radius);
    background: var(--primary);
    color: white;
    font-weight: 600;
    cursor: pointer;
    transition: opacity 0.2s ease;
}

.btn:hover {
    opacity: 0.9;
}

@media (max-width: 768px) {
    .container { padding: 0 0.5rem; }
    .btn { width: 100%; justify-content: center; }
}
```

## JSON

```json
{
    "name": "md2docx",
    "version": "0.1.5",
    "description": "Markdown to DOCX converter",
    "scripts": {
        "build": "cargo build --release",
        "test": "cargo test"
    },
    "config": {
        "languages": ["en", "th"],
        "features": {
            "syntax_highlight": true,
            "math_equations": true,
            "mermaid_diagrams": true
        },
        "limits": {
            "max_file_size_mb": 50,
            "max_images": 100
        }
    }
}
```

## YAML

```yaml
name: CI/CD Pipeline
on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        rust: [stable, nightly]
    steps:
      - uses: actions/checkout@v4
      - name: Install Rust
        uses: dtolnay/rust-toolchain@master
        with:
          toolchain: ${{ matrix.rust }}
      - name: Build
        run: cargo build --release
      - name: Test
        run: cargo test --all
      - name: Upload artifact
        if: matrix.rust == 'stable'
        uses: actions/upload-artifact@v4
        with:
          name: binary
          path: target/release/md2docx
```

## TOML

```toml
[package]
name = "my-project"
version = "1.0.0"
edition = "2021"
authors = ["Developer <dev@example.com>"]

[dependencies]
serde = { version = "1", features = ["derive"] }
tokio = { version = "1", features = ["full"] }

[dev-dependencies]
criterion = "0.5"

[[bench]]
name = "benchmarks"
harness = false

[profile.release]
opt-level = 3
lto = true
strip = true
```

## Markdown

```markdown
# Project README

> A powerful tool for document conversion.

## Features

- **Syntax Highlighting** — supports 15+ languages
- **Math Equations** — LaTeX via `$...$` and `$$...$$`
- *Mermaid Diagrams* — flowcharts, sequences, and more

| Feature    | Status |
|------------|--------|
| Highlight  | ✅     |
| Math       | ✅     |
| Mermaid    | ✅     |

See [documentation](https://example.com) for details.
```

## Plain Text (no language)

```
This is a plain code block without any language specified.
No syntax highlighting is applied here.
Lines are displayed in monospace font with the default color.
```
