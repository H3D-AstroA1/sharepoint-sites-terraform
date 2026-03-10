"""
Attachment generation for emails with department-relevant files.

Generates realistic file attachments (Word, Excel, PowerPoint, PDF)
with content appropriate for the recipient's department.
"""

import base64
import random
from datetime import datetime
from typing import Dict, Optional, List

# =============================================================================
# ATTACHMENT TEMPLATES BY DEPARTMENT
# =============================================================================

ATTACHMENT_TEMPLATES = {
    "Executive Leadership": [
        {"name": "Board_Meeting_Minutes_{date}.docx", "type": "docx"},
        {"name": "Strategic_Plan_Q{quarter}_{year}.pptx", "type": "pptx"},
        {"name": "Executive_Summary_{month}_{year}.pdf", "type": "pdf"},
        {"name": "Leadership_OKRs_{year}.xlsx", "type": "xlsx"},
        {"name": "Investor_Update_Q{quarter}.pptx", "type": "pptx"},
    ],
    "Human Resources": [
        {"name": "Employee_Handbook_v{version}.pdf", "type": "pdf"},
        {"name": "Benefits_Summary_{year}.xlsx", "type": "xlsx"},
        {"name": "Onboarding_Checklist.docx", "type": "docx"},
        {"name": "Performance_Review_Template.docx", "type": "docx"},
        {"name": "Training_Schedule_Q{quarter}.xlsx", "type": "xlsx"},
        {"name": "HR_Metrics_Report_{month}.xlsx", "type": "xlsx"},
    ],
    "Finance Department": [
        {"name": "Budget_Report_Q{quarter}_{year}.xlsx", "type": "xlsx"},
        {"name": "Financial_Statement_{month}_{year}.pdf", "type": "pdf"},
        {"name": "Expense_Report_{date}.xlsx", "type": "xlsx"},
        {"name": "Cash_Flow_Analysis_{month}.xlsx", "type": "xlsx"},
        {"name": "Audit_Findings_{year}.docx", "type": "docx"},
        {"name": "Tax_Summary_{year}.pdf", "type": "pdf"},
    ],
    "IT Department": [
        {"name": "System_Architecture_v{version}.pptx", "type": "pptx"},
        {"name": "Security_Policy_{year}.pdf", "type": "pdf"},
        {"name": "Project_Status_Report_{date}.docx", "type": "docx"},
        {"name": "IT_Roadmap_{year}.pptx", "type": "pptx"},
        {"name": "Incident_Report_{date}.docx", "type": "docx"},
        {"name": "Software_Inventory_{year}.xlsx", "type": "xlsx"},
    ],
    "Marketing Department": [
        {"name": "Campaign_Results_Q{quarter}.pptx", "type": "pptx"},
        {"name": "Brand_Guidelines_v{version}.pdf", "type": "pdf"},
        {"name": "Marketing_Plan_{year}.pptx", "type": "pptx"},
        {"name": "Social_Media_Report_{month}.xlsx", "type": "xlsx"},
        {"name": "Content_Calendar_{month}.xlsx", "type": "xlsx"},
        {"name": "Competitor_Analysis_{year}.docx", "type": "docx"},
    ],
    "Sales Department": [
        {"name": "Sales_Report_Q{quarter}_{year}.xlsx", "type": "xlsx"},
        {"name": "Client_Proposal_{company}.pptx", "type": "pptx"},
        {"name": "Contract_Template_v{version}.docx", "type": "docx"},
        {"name": "Pipeline_Analysis_{month}.xlsx", "type": "xlsx"},
        {"name": "Territory_Plan_{region}.pptx", "type": "pptx"},
        {"name": "Commission_Report_{month}.xlsx", "type": "xlsx"},
    ],
    "Legal & Compliance": [
        {"name": "NDA_Template_v{version}.docx", "type": "docx"},
        {"name": "Compliance_Checklist_{year}.xlsx", "type": "xlsx"},
        {"name": "Privacy_Policy_v{version}.pdf", "type": "pdf"},
        {"name": "Contract_Review_{date}.docx", "type": "docx"},
        {"name": "Regulatory_Update_{month}.pdf", "type": "pdf"},
    ],
    "Operations Department": [
        {"name": "Operations_Report_{month}.xlsx", "type": "xlsx"},
        {"name": "Process_Documentation_v{version}.docx", "type": "docx"},
        {"name": "Inventory_Report_{date}.xlsx", "type": "xlsx"},
        {"name": "Quality_Metrics_Q{quarter}.xlsx", "type": "xlsx"},
        {"name": "SOP_{process}.pdf", "type": "pdf"},
    ],
    "Customer Service": [
        {"name": "Support_Metrics_{month}.xlsx", "type": "xlsx"},
        {"name": "Customer_Feedback_Q{quarter}.pptx", "type": "pptx"},
        {"name": "Training_Manual_v{version}.pdf", "type": "pdf"},
        {"name": "FAQ_Document_v{version}.docx", "type": "docx"},
        {"name": "SLA_Report_{month}.xlsx", "type": "xlsx"},
    ],
    "default": [
        {"name": "Report_{date}.pdf", "type": "pdf"},
        {"name": "Presentation_{date}.pptx", "type": "pptx"},
        {"name": "Document_{date}.docx", "type": "docx"},
        {"name": "Data_{date}.xlsx", "type": "xlsx"},
    ],
}

# Content types for different file formats
CONTENT_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
}

# File icons for display
FILE_ICONS = {
    "docx": "📄",
    "xlsx": "📊",
    "pptx": "📑",
    "pdf": "📋",
}


class AttachmentGenerator:
    """Generates realistic attachments for emails."""
    
    def __init__(self):
        """Initialize the attachment generator."""
        self.templates = ATTACHMENT_TEMPLATES
        self.content_types = CONTENT_TYPES
    
    def generate(
        self,
        attachment_type: Optional[str] = None,
        department: str = "General"
    ) -> Dict:
        """
        Generate an attachment appropriate for the department.
        
        Args:
            attachment_type: Optional specific file type (docx, xlsx, pptx, pdf).
            department: Department name for context-appropriate files.
            
        Returns:
            Dictionary containing attachment data.
        """
        # Get templates for department
        templates = self.templates.get(department, self.templates["default"])
        
        # Filter by type if specified
        if attachment_type:
            templates = [t for t in templates if t["type"] == attachment_type]
            if not templates:
                templates = self.templates["default"]
        
        # Select random template
        template = random.choice(templates)
        
        # Generate filename with placeholders filled
        filename = self._fill_filename_placeholders(template["name"])
        file_type = template["type"]
        
        # Generate file content
        content = self._generate_file_content(file_type, filename, department)
        
        return {
            "name": filename,
            "content_type": self.content_types.get(file_type, "application/octet-stream"),
            "content_bytes": content,
            "content_base64": base64.b64encode(content).decode("utf-8"),
            "size": len(content),
            "size_display": self._format_size(len(content)),
            "type": file_type,
            "icon": FILE_ICONS.get(file_type, "📎"),
        }
    
    def _fill_filename_placeholders(self, filename: str) -> str:
        """Fill placeholders in filename."""
        now = datetime.now()
        
        replacements = {
            "{date}": now.strftime("%Y-%m-%d"),
            "{year}": str(now.year),
            "{month}": now.strftime("%B"),
            "{quarter}": str((now.month - 1) // 3 + 1),
            "{version}": f"{random.randint(1, 5)}.{random.randint(0, 9)}",
            "{company}": random.choice(["Contoso", "Fabrikam", "Northwind", "Acme"]),
            "{region}": random.choice(["North_America", "EMEA", "APAC", "LATAM"]),
            "{process}": random.choice(["Onboarding", "Procurement", "Expense", "Change_Request"]),
        }
        
        for placeholder, value in replacements.items():
            filename = filename.replace(placeholder, value)
        
        return filename
    
    def _generate_file_content(
        self,
        file_type: str,
        filename: str,
        department: str
    ) -> bytes:
        """Generate file content based on type."""
        generators = {
            "docx": self._create_docx,
            "xlsx": self._create_xlsx,
            "pptx": self._create_pptx,
            "pdf": self._create_pdf,
        }
        
        generator = generators.get(file_type, self._create_pdf)
        return generator(filename, department)
    
    def _create_docx(self, filename: str, department: str) -> bytes:
        """Create a minimal valid DOCX file."""
        # Minimal DOCX is a ZIP file with specific XML structure
        # This creates a very basic but valid DOCX
        import io
        import zipfile
        
        # Create in-memory ZIP
        buffer = io.BytesIO()
        
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
            zf.writestr('[Content_Types].xml', content_types)
            
            # _rels/.rels
            rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
            zf.writestr('_rels/.rels', rels)
            
            # word/document.xml
            document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>{filename}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Department: {department}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>This is a sample document created for testing purposes.</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>'''
            zf.writestr('word/document.xml', document)
        
        return buffer.getvalue()
    
    def _create_xlsx(self, filename: str, department: str) -> bytes:
        """Create a minimal valid XLSX file."""
        import io
        import zipfile
        
        buffer = io.BytesIO()
        
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
            zf.writestr('[Content_Types].xml', content_types)
            
            # _rels/.rels
            rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
            zf.writestr('_rels/.rels', rels)
            
            # xl/workbook.xml
            workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>'''
            zf.writestr('xl/workbook.xml', workbook)
            
            # xl/_rels/workbook.xml.rels
            wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
            zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
            
            # xl/worksheets/sheet1.xml
            sheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Document</t></is></c>
            <c r="B1" t="inlineStr"><is><t>{filename}</t></is></c>
        </row>
        <row r="2">
            <c r="A2" t="inlineStr"><is><t>Department</t></is></c>
            <c r="B2" t="inlineStr"><is><t>{department}</t></is></c>
        </row>
        <row r="3">
            <c r="A3" t="inlineStr"><is><t>Generated</t></is></c>
            <c r="B3" t="inlineStr"><is><t>{datetime.now().strftime("%Y-%m-%d")}</t></is></c>
        </row>
    </sheetData>
</worksheet>'''
            zf.writestr('xl/worksheets/sheet1.xml', sheet)
        
        return buffer.getvalue()
    
    def _create_pptx(self, filename: str, department: str) -> bytes:
        """Create a minimal valid PPTX file."""
        import io
        import zipfile
        
        buffer = io.BytesIO()
        
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>'''
            zf.writestr('[Content_Types].xml', content_types)
            
            # _rels/.rels
            rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''
            zf.writestr('_rels/.rels', rels)
            
            # ppt/presentation.xml
            presentation = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:sldIdLst>
        <p:sldId id="256" r:id="rId2"/>
    </p:sldIdLst>
</p:presentation>'''
            zf.writestr('ppt/presentation.xml', presentation)
            
            # ppt/_rels/presentation.xml.rels
            pres_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>'''
            zf.writestr('ppt/_rels/presentation.xml.rels', pres_rels)
            
            # ppt/slides/slide1.xml
            slide = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title"/>
                    <p:cNvSpPr/>
                    <p:nvPr/>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                    <a:bodyPr/>
                    <a:p>
                        <a:r>
                            <a:t>{filename}</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:r>
                            <a:t>{department}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>'''
            zf.writestr('ppt/slides/slide1.xml', slide)
        
        return buffer.getvalue()
    
    def _create_pdf(self, filename: str, department: str) -> bytes:
        """Create a minimal valid PDF file."""
        # Create a simple PDF with basic content
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # PDF content
        content = f"""Document: {filename}
Department: {department}
Generated: {now}

This is a sample document created for testing purposes.
It simulates realistic organizational content.

---
This document is confidential and intended for internal use only.
"""
        
        # Build minimal PDF structure
        pdf_content = f"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length {len(content) + 50} >>
stream
BT
/F1 12 Tf
50 750 Td
({filename}) Tj
0 -20 Td
(Department: {department}) Tj
0 -20 Td
(Generated: {now}) Tj
0 -40 Td
(This is a sample document for testing.) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000266 00000 n 
0000000{350 + len(content):03d} 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
{420 + len(content)}
%%EOF"""
        
        return pdf_content.encode('latin-1')
    
    def _format_size(self, size_bytes: int) -> str:
        """Format file size for display."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes // 1024} KB"
        else:
            return f"{size_bytes // (1024 * 1024)} MB"
