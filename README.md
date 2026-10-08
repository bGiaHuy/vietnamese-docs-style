# Vietnamese Docs Style

**Chỉnh sửa phong cách tạo docx của AI cho phù hợp với yêu cầu về văn bản của Việt Nam.**

*Localizing AI-powered Word document generation for Vietnam.*

**Ngôn ngữ / Language:** [Tiếng Việt](#tiếng-việt) | [English](#english)

---

# Tiếng Việt

## Giới thiệu

**AI có thể tạo tài liệu Word, nhưng liệu AI có thực sự hiểu cách người Việt Nam trình bày văn bản?**

`vietnamese-docs-style` là một AI Agent Skill hỗ trợ tạo, định dạng, biên tập và kiểm tra tài liệu Word (`.docx`) theo các quy ước trình bày, cấu trúc và văn phong phù hợp với bối cảnh sử dụng tại Việt Nam.

Dự án được xây dựng nhằm thu hẹp khoảng cách giữa khả năng tạo tài liệu của AI và những yêu cầu thực tế trong môi trường học thuật, hành chính và tổ chức tại Việt Nam.

## 1. Vấn đề đặt ra

Sự phát triển của các mô hình ngôn ngữ lớn (LLM) và AI Agent đã giúp việc tạo tài liệu trở nên dễ dàng hơn. Người dùng có thể yêu cầu AI viết một báo cáo, soạn công văn hoặc xuất trực tiếp nội dung thành file DOCX.

Tuy nhiên, **khả năng tạo ra một file Word hợp lệ không đồng nghĩa với khả năng tạo ra một tài liệu được trình bày đúng quy cách.**

Các mẫu định dạng phổ thông mà AI sử dụng có thể phù hợp với một số quy ước quốc tế, nhưng chưa chắc đáp ứng yêu cầu tại Việt Nam.

Chẳng hạn:

- **Văn bản hành chính** có những quy định về thể thức và kỹ thuật trình bày theo Nghị định 30/2020/NĐ-CP.
- **Báo cáo học thuật** thường phải tuân theo hướng dẫn riêng của từng trường đại học hoặc cơ sở đào tạo.
- **Biên bản và đề xuất dự án** có bố cục, cách diễn đạt và mức độ trang trọng phụ thuộc vào mục đích sử dụng.
- **Tài liệu nội bộ** có thể cần tuân theo những mẫu được tổ chức quy định sẵn.

Sự khác biệt không chỉ nằm ở phông chữ, cỡ chữ hay căn lề, mà còn ở cấu trúc nội dung, hệ thống đề mục, cách trình bày bảng biểu và văn phong.

Điều này đặt ra một câu hỏi:

**Làm thế nào để AI không chỉ viết tài liệu bằng tiếng Việt, mà còn hiểu cách trình bày tài liệu phù hợp với từng bối cảnh tại Việt Nam?**

## 2. Quá trình nghiên cứu và phát triển

`vietnamese-docs-style` được phát triển từ nhu cầu giải quyết những bất cập trong quá trình sử dụng AI để tạo tài liệu tiếng Việt.

Thay vì xây dựng một template cố định, dự án tiếp cận vấn đề bằng cách nghiên cứu và hệ thống hóa các quy tắc trình bày, cấu trúc và biên tập tài liệu.

Những nội dung được tập trung bao gồm:

- Các quy định về thể thức văn bản hành chính Việt Nam.
- Quy ước trình bày báo cáo học thuật và tài liệu nghiên cứu.
- Cấu trúc của những loại biên bản, kế hoạch và đề xuất.
- Cách sử dụng tiếng Việt trong văn bản trang trọng.
- Những khác biệt giữa yêu cầu định dạng chung và yêu cầu riêng của từng tổ chức.

Các kết quả được tổ chức thành hệ thống tài liệu tham chiếu (`references`), hướng dẫn dành cho AI Agent và những công cụ hỗ trợ xử lý DOCX.

Thay vì buộc người dùng phải giải thích lại toàn bộ quy tắc trình bày trong mỗi lần yêu cầu, skill cung cấp một cơ sở hướng dẫn có thể tái sử dụng.

## 3. Giải pháp

Vietnamese Docs Style hoạt động như một lớp tri thức chuyên biệt, giúp AI Agent lựa chọn cách tạo tài liệu theo ngữ cảnh.

### Phân loại tài liệu

Skill xác định nhóm tài liệu trước khi lựa chọn những quy tắc định dạng và cấu trúc phù hợp.

| Profile | Nhóm tài liệu |
|---|---|
| `administrative` | Văn bản hành chính |
| `academic` | Báo cáo học thuật, tiểu luận, nghiên cứu |
| `proposal` | Đề xuất dự án, kế hoạch |
| `minutes-administrative` | Biên bản hành chính |
| `minutes-general` | Biên bản họp thông thường |
| `custom` | Tài liệu theo mẫu riêng |

### Áp dụng quy tắc theo ngữ cảnh

Một nguyên tắc quan trọng của dự án là **không áp dụng một tiêu chuẩn duy nhất cho tất cả tài liệu tiếng Việt**.

Chẳng hạn, Nghị định 30/2020/NĐ-CP được sử dụng làm cơ sở tham chiếu cho những văn bản hành chính thuộc phạm vi áp dụng, thay vì áp đặt lên mọi báo cáo học thuật hay tài liệu doanh nghiệp.

Khi người dùng cung cấp mẫu hoặc yêu cầu riêng, AI Agent được hướng dẫn xem xét những yêu cầu đó cùng các quy định liên quan.

### Hỗ trợ tạo và kiểm tra DOCX

Dự án cung cấp các công cụ Python nhằm hỗ trợ:

- Thiết lập khổ giấy, căn lề, phông chữ và định dạng cơ bản.
- Xây dựng đoạn văn, tiêu đề, bảng và một số thành phần tài liệu.
- Kiểm tra một số thuộc tính cấu trúc và định dạng.
- Phát hiện những placeholder chưa được xử lý.

## 4. Quy trình hoạt động

**Bước 1 — Phân tích yêu cầu**

Xác định loại tài liệu, mục đích sử dụng, thông tin đầu vào và những yêu cầu định dạng đặc biệt.

**Bước 2 — Lựa chọn Document Profile**

Chọn nhóm tài liệu tương ứng và xác định bộ quy tắc cần sử dụng.

**Bước 3 — Tham chiếu quy chuẩn**

AI Agent đọc các tài liệu trong `references/` để lựa chọn cách trình bày, cấu trúc và văn phong thích hợp.

**Bước 4 — Tạo tài liệu**

Sử dụng hướng dẫn và các công cụ hỗ trợ để xây dựng file DOCX.

**Bước 5 — Kiểm tra**

Thực hiện các bước kiểm tra được hỗ trợ, đồng thời rà soát kết quả hiển thị khi có công cụ render phù hợp.

## 5. Cấu trúc dự án

```text
vietnamese-docs-style/
├── README.md
├── SKILL.md
├── GEMINI.md
├── references/
│   ├── administrative-format-nd30.md
│   ├── academic-report.md
│   ├── document-profiles.md
│   ├── editorial-quality-vi.md
│   ├── meeting-minutes.md
│   ├── style-spec.md
│   └── validation-checklist.md
├── scripts/
│   ├── build_docx.py
│   ├── validate_docx.py
│   └── render_docx.py
├── assets/
└── tests/
```

Trong đó:

- `SKILL.md`: Điểm vào chính và hướng dẫn quy trình dành cho AI Agent.
- `references/`: Tài liệu tham chiếu về cấu trúc, thể thức và ngôn ngữ.
- `scripts/`: Công cụ hỗ trợ tạo và kiểm tra DOCX.
- `assets/`: Tài nguyên và tài liệu mẫu.
- `tests/`: Các bài kiểm thử.

## 6. Sử dụng

Skill được thiết kế để tích hợp vào môi trường AI Agent có hỗ trợ đọc hướng dẫn và thực thi công cụ phù hợp.

Người dùng có thể tham khảo `SKILL.md` để cấu hình skill trong môi trường agent của mình.

### Ví dụ yêu cầu

**Văn bản hành chính**

> Tạo một công văn thông báo lịch họp, áp dụng quy định trình bày phù hợp theo Nghị định 30/2020/NĐ-CP.

**Báo cáo học thuật**

> Viết báo cáo nghiên cứu về ứng dụng trí tuệ nhân tạo trong giáo dục, có trang bìa và hệ thống đề mục rõ ràng.

**Biên bản họp**

> Tạo biên bản họp nhóm từ những ghi chú này, bao gồm nội dung thảo luận, kết luận và phân công công việc.

**Tài liệu tùy chỉnh**

> Định dạng lại tài liệu Word theo mẫu do trường đại học cung cấp.

### Kiểm tra DOCX

Có thể sử dụng công cụ kiểm tra cấu trúc được cung cấp trong repository:

```bash
python scripts/validate_docx.py document.docx --profile administrative
```

Cần cài đặt các thư viện phụ thuộc phù hợp, bao gồm `python-docx`.

## 7. Giới hạn

Vietnamese Docs Style là một công cụ hỗ trợ AI Agent, không phải hệ thống chứng nhận tính hợp lệ của văn bản.

- Kết quả phụ thuộc vào khả năng của AI Agent và thông tin được cung cấp.
- Các công cụ kiểm tra hiện tại chưa bao phủ toàn bộ quy định về thể thức và trình bày.
- Một số vấn đề về bố cục cần được kiểm tra trực quan.
- Những tài liệu sử dụng trong hoạt động hành chính chính thức vẫn cần được người có trách nhiệm rà soát.

## 8. Định hướng

Mục tiêu của dự án là phát triển một hệ thống tri thức và công cụ hỗ trợ AI tạo tài liệu tiếng Việt một cách nhất quán, có ngữ cảnh và phù hợp hơn với môi trường sử dụng thực tế.

Bản địa hóa AI không chỉ là chuyển đổi ngôn ngữ. Đó còn là giúp AI hiểu các quy ước, tiêu chuẩn và cách thức làm việc của cộng đồng mà nó phục vụ.

---

# English

## Introduction

**AI can generate Word documents. But does it really understand how documents are written and formatted in Vietnam?**

`vietnamese-docs-style` is an AI Agent Skill designed to assist with creating, formatting, editing, and validating Microsoft Word (`.docx`) documents according to Vietnamese document conventions, structures, and writing practices.

The project aims to bridge the gap between general-purpose AI document generation and the specific requirements found in Vietnamese academic, administrative, and organizational environments.

## 1. The Problem

Large Language Models (LLMs) and AI agents have made document creation significantly more accessible.

Users can now ask an AI to write a report, prepare an official letter, or generate a complete Word document.

However, **generating a valid DOCX file does not necessarily mean producing a properly formatted document.**

General-purpose AI tools may rely on document templates and formatting conventions that work well in some international contexts but do not necessarily match Vietnamese requirements.

For example:

- **Administrative documents** in Vietnam are subject to specific presentation and formatting requirements under Decree 30/2020/NĐ-CP.
- **Academic reports** often follow institution-specific formatting and structural guidelines.
- **Meeting minutes and project proposals** require different layouts, terminology, and levels of formality depending on their purpose.
- **Internal documents** may need to follow organization-specific templates.

These differences extend beyond fonts, margins, and page sizes. They also involve document structure, heading hierarchy, tables, and writing conventions.

This raises an important question:

**How can AI move beyond generating Vietnamese text and produce documents that respect the conventions of their intended local context?**

## 2. Research and Development

Vietnamese Docs Style was developed to address practical limitations encountered when using AI to create Vietnamese Word documents.

Instead of building another fixed document template, the project focuses on researching and systematizing local document conventions into reusable guidance.

The research focuses on:

- Vietnamese administrative document formatting requirements.
- Academic and research document presentation conventions.
- Structures used in meeting minutes, plans, and project proposals.
- Vietnamese editorial practices and formal writing styles.
- Differences between general formatting standards and institution-specific requirements.

This knowledge is organized into reference documents, agent instructions, and supporting Python utilities.

The goal is to reduce the need for users to repeatedly explain the same formatting requirements each time they generate a document.

## 3. The Solution

Vietnamese Docs Style acts as a specialized knowledge layer that helps AI agents make context-aware document generation decisions.

### Document classification

The skill identifies the document category before selecting the appropriate formatting and structural guidance.

| Profile | Intended use |
|---|---|
| `administrative` | Administrative documents |
| `academic` | Academic reports, essays, research papers |
| `proposal` | Project proposals and planning documents |
| `minutes-administrative` | Administrative meeting minutes |
| `minutes-general` | General and internal meeting minutes |
| `custom` | User-provided templates and requirements |

### Context-aware formatting

A key design principle is that **no single formatting standard should be applied to every Vietnamese document**.

For example, Decree 30/2020/NĐ-CP is referenced when preparing applicable administrative documents, rather than being automatically imposed on academic reports or business documents.

When users provide their own templates or organizational requirements, the agent is instructed to consider them alongside relevant standards.

### DOCX generation and validation support

The repository includes Python utilities that assist with:

- Configuring page dimensions, margins, fonts, and basic styles.
- Creating paragraphs, headings, tables, and document elements.
- Checking selected structural and formatting properties.
- Detecting unresolved placeholders.

## 4. Workflow

**Step 1 — Understand the request**

Identify the document type, intended purpose, source information, and any specific formatting requirements.

**Step 2 — Select a document profile**

Choose the appropriate document category and determine which conventions should apply.

**Step 3 — Consult references**

Read relevant guidance from `references/` to determine document structure, formatting, and writing conventions.

**Step 4 — Generate the document**

Use the available instructions and utilities to construct the DOCX output.

**Step 5 — Validate the result**

Run supported checks and visually review the rendered document when an appropriate rendering tool is available.

## 5. Repository Structure

```text
vietnamese-docs-style/
├── README.md
├── SKILL.md
├── GEMINI.md
├── references/
│   ├── administrative-format-nd30.md
│   ├── academic-report.md
│   ├── document-profiles.md
│   ├── editorial-quality-vi.md
│   ├── meeting-minutes.md
│   ├── style-spec.md
│   └── validation-checklist.md
├── scripts/
│   ├── build_docx.py
│   ├── validate_docx.py
│   └── render_docx.py
├── assets/
└── tests/
```

- `SKILL.md`: Main entry point and workflow instructions for AI agents.
- `references/`: Document formatting, structural, and editorial knowledge.
- `scripts/`: Supporting DOCX generation and validation utilities.
- `assets/`: Document samples and related resources.
- `tests/`: Test files.

## 6. Usage

The skill is intended for AI agent environments capable of reading skill instructions and using compatible tools.

Refer to `SKILL.md` for the agent workflow and relevant integration details.

### Example prompts

**Administrative document**

> Create a Vietnamese administrative notice using the applicable formatting guidance from Decree 30/2020/NĐ-CP.

**Academic report**

> Prepare an academic report on artificial intelligence in education, including a cover page and structured headings.

**Meeting minutes**

> Convert these meeting notes into formal Vietnamese meeting minutes, including decisions and assigned responsibilities.

**Custom formatting**

> Reformat this Word document according to the university template I provided.

### DOCX validation

The repository includes a structural validation utility:

```bash
python scripts/validate_docx.py document.docx --profile administrative
```

Compatible dependencies, including `python-docx`, must be installed.

## 7. Limitations

Vietnamese Docs Style assists AI agents with document preparation. It does not certify that documents are legally compliant or universally correct.

- Output quality depends on the AI agent and information provided.
- Current validation utilities cover only selected formatting and structural checks.
- Some layout problems require visual inspection.
- Official administrative documents still require appropriate human review.

## 8. Project Vision

The long-term goal is to build a reusable knowledge and tooling foundation that helps AI agents generate Vietnamese documents more consistently and with greater awareness of local requirements.

**AI localization is more than language translation. It is about understanding the conventions, standards, and workflows of the people who use it.**

---

## Tác giả / Author

**Bùi Gia Huy**

GitHub: [bGiaHuy](https://github.com/bGiaHuy)

Repository: [vietnamese-docs-style](https://github.com/bGiaHuy/vietnamese-docs-style)
