# Slide Translator - Data Flow Diagram

## System Overview: English → Arabic Translation Pipeline

```mermaid
flowchart TB
    Start([English PowerPoint File]) --> Extract[Stage 1: Extraction Layer]

    Extract --> ExtractSub{Decompose Archive}
    ExtractSub --> Slides[Slide Definitions]
    ExtractSub --> Layouts[Layout Templates]
    ExtractSub --> Charts[Chart Data Files]
    ExtractSub --> Masters[Master Styles]

    Slides --> Content[Stage 2: Content Processing]
    Layouts --> Content
    Charts --> ChartProc[Chart Text Extraction]
    Masters --> Content

    Content --> Analysis{Analyze Each Element}
    Analysis --> Title[Identify Titles]
    Analysis --> Body[Identify Body Text]
    Analysis --> Tables[Identify Table Content]
    Analysis --> Anno[Identify Annotations]

    Title --> Package[Package with Context]
    Body --> Package
    Tables --> Package
    Anno --> Package

    Package --> JSON1[Structured JSON Document]
    ChartProc --> JSON2[Chart Text Document]

    JSON1 --> AI[Stage 3: AI Translation Engine]
    JSON2 --> AI

    AI --> Context{Apply Context Understanding}
    Context --> Translate[Translate to Arabic]
    Translate --> Tone[Maintain Professional Tone]
    Tone --> JSON3[Translated JSON Document]

    JSON3 --> Parallel{Parallel Processing}

    Parallel --> Visual[Stage 4A: Visual Transformation]
    Parallel --> Inject[Stage 4B: Text Injection]

    Visual --> Mirror[Calculate Mirror Positions]
    Mirror --> Formula["Formula: New_X = Width - Old_X - Element_Width"]
    Formula --> Direction[Reverse Text Direction]
    Direction --> TableRev[Reverse Table Columns]
    TableRev --> ChartFlip[Flip Bar Charts]
    ChartFlip --> Fonts[Inject Arabic Fonts]

    Inject --> Match[Match Translated Text to Elements]
    Match --> Apply[Apply Translations]
    Apply --> Format[Preserve Formatting]

    Fonts --> Merge[Stage 5: Merge Results]
    Format --> Merge

    Merge --> Modified[Modified Slide Files]
    Modified --> Rebuild[Stage 6: Reconstruction]

    Rebuild --> Assemble[Reassemble Archive]
    Assemble --> Validate[Validate Output]
    Validate --> Output([Arabic PowerPoint File])

    style Start fill:#e1f5ff
    style Output fill:#c8e6c9
    style AI fill:#fff9c4
    style Visual fill:#f8bbd0
    style Inject fill:#f8bbd0
    style Extract fill:#d1c4e9
    style Rebuild fill:#d1c4e9
```

---

## Processing Stages Explained

### Stage 1: Extraction
**Input**: Single PowerPoint file
**Process**: Decompress and separate into components
**Output**: Individual data files (slides, charts, layouts)

### Stage 2: Content Processing
**Input**: Slide data files
**Process**: Identify all text and determine its role (title/body/annotation)
**Output**: Structured document with text + context

### Stage 3: AI Translation
**Input**: Text with context ("This is a title in a strategy slide")
**Process**: AI translates maintaining consulting terminology
**Output**: Arabic translations in same structure

### Stage 4A: Visual Transformation
**Input**: Original slide layout
**Process**: Mirror every element mathematically for RTL
**Output**: Reversed visual layout

### Stage 4B: Text Injection
**Input**: Translated text + transformed layout
**Process**: Insert Arabic text into mirrored positions
**Output**: Complete Arabic slides

### Stage 5-6: Reconstruction
**Input**: All transformed components
**Process**: Reassemble into PowerPoint format
**Output**: Final Arabic presentation

---

## Key Data Transformations

```mermaid
graph LR
    A[Text in English] --> B[Text + Context]
    B --> C[AI Translation]
    C --> D[Arabic Text]
    D --> E[RTL Layout]
    E --> F[Final Slide]

    style A fill:#bbdefb
    style D fill:#c8e6c9
    style F fill:#a5d6a7
```

---

## Parallel Processing Architecture

```mermaid
graph TB
    Input[Translated Content] --> Split{Split Processing}

    Split -->|Path A| Visual[Visual Transformation<br/>Mirror positions<br/>Reverse directions<br/>Flip charts]
    Split -->|Path B| Text[Text Injection<br/>Apply translations<br/>Set fonts<br/>Preserve formatting]

    Visual --> Combine[Combine Results]
    Text --> Combine

    Combine --> Complete[Complete Arabic Slide]

    style Split fill:#fff9c4
    style Combine fill:#c8e6c9
```

---

## Chart Processing Flow (Advanced Feature)

```mermaid
flowchart LR
    Chart[Chart File] --> Extract[Extract Text]
    Extract --> Title[Chart Title]
    Extract --> Series[Series Names]
    Extract --> Categories[Axis Labels]

    Title --> Translate[AI Translation]
    Series --> Translate
    Categories --> Translate

    Translate --> Inject[Inject Translations]
    Inject --> Detect{Is Horizontal<br/>Bar Chart?}

    Detect -->|Yes| Flip[Reverse Axis Orientation<br/>Bars grow Right→Left]
    Detect -->|No| Keep[Keep Original Orientation]

    Flip --> Final[Arabic Chart]
    Keep --> Final

    style Detect fill:#fff9c4
    style Flip fill:#f8bbd0
```

---

## Quality Assurance Checkpoints

```mermaid
flowchart TD
    Start([Processing Start]) --> Check1{Input Valid?}
    Check1 -->|No| Error1[Report Error & Stop]
    Check1 -->|Yes| Process1[Extract Components]

    Process1 --> Check2{All Text<br/>Extracted?}
    Check2 -->|No| Log1[Log Issue<br/>Continue with Available]
    Check2 -->|Yes| Process2[Translate Content]
    Log1 --> Process2

    Process2 --> Check3{Translation<br/>Complete?}
    Check3 -->|No| Retry[Retry Translation]
    Check3 -->|Yes| Process3[Apply Transformations]
    Retry --> Check3

    Process3 --> Check4{Visual Integrity<br/>Maintained?}
    Check4 -->|No| Log2[Log Issue<br/>Use Best Available]
    Check4 -->|Yes| Process4[Rebuild File]
    Log2 --> Process4

    Process4 --> Check5{Output Opens<br/>in PowerPoint?}
    Check5 -->|No| Error2[Critical Error<br/>Manual Review]
    Check5 -->|Yes| Success([Deliver to Client])

    style Check1 fill:#fff9c4
    style Check2 fill:#fff9c4
    style Check3 fill:#fff9c4
    style Check4 fill:#fff9c4
    style Check5 fill:#fff9c4
    style Success fill:#c8e6c9
    style Error1 fill:#ffcdd2
    style Error2 fill:#ffcdd2
```
