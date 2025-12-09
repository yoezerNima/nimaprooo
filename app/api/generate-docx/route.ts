import { Document, Packer, Paragraph, HeadingLevel, AlignmentType, PageBreak, TextRun } from "docx"

interface DocumentInput {
  title: string
  problemStatement?: string
  objectives?: string[]
  requirements?: string[]
  manualSteps?: string[]
  automationIdeas?: string
}

export async function POST(request: Request) {
  try {
    const body: DocumentInput = await request.json()

    // Validate input
    if (!body.title || typeof body.title !== "string") {
      return Response.json({ error: "Invalid input: title is required and must be a string" }, { status: 400 })
    }

    if (body.objectives !== undefined && !Array.isArray(body.objectives)) {
      return Response.json(
        { error: "Invalid input: objectives must be an array" },
        { status: 400 },
      )
    }

    if (body.requirements !== undefined && !Array.isArray(body.requirements)) {
      return Response.json(
        { error: "Invalid input: requirements must be an array" },
        { status: 400 },
      )
    }

    if (body.manualSteps !== undefined && !Array.isArray(body.manualSteps)) {
      return Response.json(
        { error: "Invalid input: manualSteps must be an array" },
        { status: 400 },
      )
    }

    if (body.automationIdeas !== undefined && typeof body.automationIdeas !== "string") {
      return Response.json(
        { error: "Invalid input: automationIdeas must be a string" },
        { status: 400 },
      )
    }

    if (body.problemStatement !== undefined && typeof body.problemStatement !== "string") {
      return Response.json(
        { error: "Invalid input: problemStatement must be a string" },
        { status: 400 },
      )
    }

    // Create sections following PDD template structure
    const sections = [
      // Header with footer text
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
            size: 36,
          }),
        ],
      }),

      // Title section
      new Paragraph({
        children: [
          new TextRun({
            text: `Process: TMPY HR`,
            size: 40,
          }),
        ],
        spacing: { after: 100 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Project Name: ${body.title}`,
            size: 40,
          }),
        ],
        spacing: { after: 100 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Nov | 2024`,
            size: 40,
          }),
        ],
        spacing: { after: 600 },
      }),

      new Paragraph({
        children: [
          new TextRun({
            text: "Process Definition Document",
            bold: true,
            size: 56,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),

      // Footer line
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        pageBreakBefore: true,
        children: [
          new TextRun({
            text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
            size: 36,
          }),
        ],
      }),

      // Table of Contents
      new Paragraph({
        children: [
          new TextRun({
            text: "1 Contents",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { after: 200 },
      }),
      createTOCItem("2 INTRODUCTION", 3),
      createTOCItem("2.1 PURPOSE", 3),
      createTOCItem("2.2 OBJECTIVE", 3),
      createTOCItem("2.3 PROCESS KEY CONTACTS", 3),
      createTOCItem("2.4 DOCUMENT CONTROL", 3),
      createTOCItem("3 CHANGE REQUESTS", 3),
      createTOCItem("4 AS IS PROCESS DESCRIPTION", 3),
      createTOCItem("4.1 PROCESS OVERVIEW", 4),
      createTOCItem("4.2 RACI MATRIX", 4),
      createTOCItem("4.3 MINIMUM PRE-REQUISITES FOR THE AUTOMATION", 4),
      createTOCItem("4.4 APPLICATION USED IN THE PROCESS", 4),
      createTOCItem("4.5 AS-IS PROCESS MAP", 4),
      createTOCItem("4.5.1 High Level AS-IS Process Map", 5),
      createTOCItem("4.5.2 Detailed AS-IS Process Map", 5),
      createTOCItem("4.6 VOLUMETRIC", 4),
      createTOCItem("4.7 VOLUME AND HANDLING TIME", 4),
      createTOCItem("4.8 OPERATING WINDOW & STAFFING SCHEDULE", 4),
      createTOCItem("4.9 INPUT DATA DETAILS", 4),
      createTOCItem("5 TO BE PROCESS DESCRIPTION", 3),
      createTOCItem("5.1 TO BE DETAILED PROCESS MAP", 4),
      createTOCItem("5.2 PARALLEL INITIATIVES/ AUTOMATION/ DEVELOPMENT", 4),
      createTOCItem("5.3 IN SCOPE SCENARIOS/CASE TYPES/VOLUME", 4),
      createTOCItem("5.4 OUT OF SCOPE SCENARIOS/CASE TYPES/VOLUME FOR PROJECT", 4),
      createTOCItem("5.5 EXCEPTION HANDLING", 4),
      createTOCItem("5.5.1 Known Business Exception", 5),
      createTOCItem("5.5.2 Unknown Business Exception", 5),
      createTOCItem("5.6 APPLICATIONS ERRORS & EXCEPTIONS HANDLING", 4),
      createTOCItem("5.6.1 Known Applications Errors and Exceptions", 5),
      createTOCItem("5.6.2 Unknown Applications Errors and Exceptions", 5),
      createTOCItem("5.7 REPORTING", 4),
      createTOCItem("6 OTHER", 3),
      createTOCItem("6.1 APPENDIX & OTHER DOCUMENTS", 4),

      new Paragraph({
        text: "",
        spacing: { after: 400 },
      }),

      // Page break before main content
      new PageBreak(),

      // Footer line
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
            size: 36,
          }),
        ],
      }),

      // INTRODUCTION section
      new Paragraph({
        children: [
          new TextRun({
            text: "2 INTRODUCTION",
            bold: true,
            size: 52,
          }),
        ],
        spacing: { before: 400, after: 300 },
      }),

      // Problem Statement
      new Paragraph({
        children: [
          new TextRun({
            text: "1. Problem Statement",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { before: 200, after: 150 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: body.problemStatement || "No problem statement provided.",
            size: 44,
          }),
        ],
        spacing: { after: 300 },
      }),

      // Objectives
      new Paragraph({
        children: [
          new TextRun({
            text: "2. Objectives",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { before: 200, after: 150 },
      }),
      ...createBulletList(body.objectives || []),

      // Requirements
      new Paragraph({
        children: [
          new TextRun({
            text: "3. Requirements",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { before: 200, after: 150 },
      }),
      ...createBulletList(body.requirements || []),

      // AS-IS Process Map
      new Paragraph({
        children: [
          new TextRun({
            text: "4. AS-IS Process Map",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { before: 200, after: 150 },
      }),
      ...createNumberedList(body.manualSteps || []),

      // Automation Ideas
      new Paragraph({
        children: [
          new TextRun({
            text: "5. Automation Ideas",
            bold: true,
            size: 44,
          }),
        ],
        spacing: { before: 200, after: 150 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: body.automationIdeas || "No automation ideas provided.",
            size: 44,
          }),
        ],
        spacing: { after: 600 },
      }),

      // Footer on last section
      new Paragraph({
        text: "",
        spacing: { after: 400 },
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "Created & Managed by Team, for queries write us at ABC@XYZ.COM",
            size: 36,
          }),
        ],
      }),
    ]

    // Create document
    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: sections,
        },
      ],
    })

    // Generate buffer and convert to base64
    const buffer = await Packer.toBuffer(doc)
    const base64 = buffer.toString("base64")

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, -5)
    const fileName = `document-${timestamp}.docx`

    return Response.json(
      {
        fileName,
        fileBase64: base64,
      },
      { status: 200 },
    )
  } catch (error) {
    console.error("Error generating DOCX:", error)
    return Response.json(
      { error: "Failed to generate document. Please check your input and try again." },
      { status: 500 },
    )
  }
}

// Helper function to create bullet list items
function createBulletList(items: string[]): Paragraph[] {
  return items
    .filter((item) => item && item.trim())
    .map(
      (item) =>
        new Paragraph({
          children: [
            new TextRun({
              text: item,
              size: 44,
            }),
          ],
          bullet: {
            level: 0,
          },
          spacing: { after: 100 },
        }),
    )
}

// Helper function to create numbered list items
function createNumberedList(items: string[]): Paragraph[] {
  return items
    .filter((item) => item && item.trim())
    .map(
      (item, index) =>
        new Paragraph({
          children: [
            new TextRun({
              text: item,
              size: 44,
            }),
          ],
          numbering: {
            reference: "num",
            level: 0,
            instance: index,
          },
          spacing: { after: 100 },
        }),
    )
}

// Helper function for table of contents formatting
function createTOCItem(text: string, indentLevel: number): Paragraph {
  return new Paragraph({
    children: [
      new TextRun({
        text: text,
        size: 40,
      }),
    ],
    spacing: { after: 100 },
    indent: { left: (indentLevel - 1) * 400 },
  })
}
