import os
import tempfile
import re
from datetime import datetime
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, after_this_request
import pdfplumber
from groq import Groq
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, gray

# Load environment variables for local development
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ================= CONFIGURATION =================
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GROQ_MODEL = "llama-3.3-70b-versatile"

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "hr-resume-converter-2026")
TEMP_DIR = tempfile.gettempdir()

# ================= HELPERS =================
def extract_text(path):
    """Extract text from PDF or DOCX"""
    text = ""
    try:
        if path.lower().endswith(".pdf"):
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or "") + "\n"
        else:
            # Fallback for DOCX text extraction without python-docx
            import zipfile
            import xml.etree.ElementTree as ET
            with zipfile.ZipFile(path) as docx:
                xml_content = docx.read('word/document.xml')
                tree = ET.fromstring(xml_content)
                namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                for paragraph in tree.findall('.//w:p', namespace):
                    texts = [node.text for node in paragraph.findall('.//w:t', namespace) if node.text]
                    if texts:
                        text += "".join(texts) + "\n"
    except Exception as e:
        print(f"Extraction Error: {e}")
    return text.strip()

def get_ai_data(resume_text):
    """Get structured data from Groq AI"""
    if not GROQ_API_KEY:
        print("API Key missing!")
        return None

    prompt = f"""Extract professional details from this resume.
STRICT RULES:
1. ONLY use details from the text.
2. If missing, leave empty.
3. Separate Technical and Soft skills.
4. Professional Title should be JUST the job title (e.g. "Java Developer").

FORMAT:
CANDIDATE INFORMATION:
- Full Name:
- Professional Title:
- Email:
- Phone:
- Location:

PROFILE SUMMARY:
(Short summary)

PROFESSIONAL EXPERIENCE:
(Format: Company | Role | Duration | Responsibilities)

PROJECT EXPERIENCE:
(Projects and tech used)

TECHNICAL SKILLS:
(Tools, languages, etc.)

SOFT SKILLS:
(Communication, etc.)

TEXT:
{resume_text[:4000]}
"""
    try:
        client = Groq(api_key=GROQ_API_KEY)
        resp = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2
        )
        return parse_ai_response(resp.choices[0].message.content)
    except Exception as e:
        print(f"AI Error: {e}")
        return None

def parse_ai_response(text):
    """Parse AI output into dictionary"""
    data = {
        "Full Name": "", "Professional Title": "", "Email": "", "Phone": "", "Location": "",
        "Profile Summary": "", "Professional Experience": "", "Project Experience": "",
        "Technical Skills": "", "Soft Skills": ""
    }
    
    current_key = None
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        
        lower_line = line.lower()
        if "full name:" in lower_line: data["Full Name"] = line.split(":", 1)[1].strip(); continue
        if "professional title:" in lower_line: data["Professional Title"] = line.split(":", 1)[1].strip(); continue
        if "email:" in lower_line: data["Email"] = line.split(":", 1)[1].strip(); continue
        if "phone:" in lower_line: data["Phone"] = line.split(":", 1)[1].strip(); continue
        if "location:" in lower_line: data["Location"] = line.split(":", 1)[1].strip(); continue
        
        sec_map = {
            "profile summary:": "Profile Summary",
            "professional experience:": "Professional Experience",
            "project experience:": "Project Experience",
            "technical skills:": "Technical Skills",
            "soft skills:": "Soft Skills"
        }
        for k, v in sec_map.items():
            if k in lower_line:
                current_key = v
                break
        else:
            if current_key:
                data[current_key] += line + "\n"
            
    return {k: v.strip() for k, v in data.items()}

def create_resume_pdf(data, path):
    """Generate PDF directly using ReportLab (Cross-Platform)"""
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    margin = 0.5 * inch
    
    # 1. Page Border (Blue)
    c.setStrokeColor(HexColor('#2980B9'))
    c.setLineWidth(1)
    c.rect(margin, margin, width - 2*margin, height - 2*margin)
    
    y = height - margin - 0.4*inch
    
    # 2. Logo (Centered)
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    logo_path = None
    for ext in ['png', 'jpg', 'jpeg']:
        lp = os.path.join(static_dir, f'krify_logo.{ext}')
        if os.path.exists(lp):
            logo_path = lp
            break
            
    if logo_path:
        logo_w, logo_h = 1.2*inch, 0.6*inch
        c.drawImage(logo_path, (width - logo_w)/2, y - logo_h, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
        y -= logo_h + 0.3*inch
    else:
        c.setFont("Helvetica-Bold", 18)
        c.setFillColor(HexColor('#003765'))
        c.drawCentredString(width/2, y, "KRIFY")
        y -= 0.5*inch

    # 3. Name & Title
    name = data["Full Name"].upper()
    title = re.sub(r'^(PROFESSIONAL TITLE/DESIGNATION:|TITLE:|DESIGNATION:)\s*', '', data["Professional Title"], flags=re.I).strip().upper()
    display = f"{name} - {title}" if title else name
    
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(HexColor('#003765'))
    c.drawCentredString(width/2, y, display)
    y -= 0.3*inch

    # 4. Contact
    parts = [v for k, v in data.items() if k in ["Email", "Phone", "Location"] and v]
    if parts:
        c.setFont("Helvetica", 9)
        c.setFillColor(gray)
        c.drawCentredString(width/2, y, " | ".join(parts))
        y -= 0.5*inch

    # Helper for adding wrapped sections
    def add_section(title, content):
        nonlocal y
        if not content: return
        
        # INCREASED Space before section (gap between sections)
        y -= 0.45*inch

        # Section Title
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(HexColor('#003765'))
        c.drawString(margin + 0.2*inch, y, title)
        y -= 0.25*inch
        
        # Section Content (Simple Wrap)
        c.setFont("Helvetica", 10)
        c.setFillColor(HexColor('#34495E'))
        
        lines = content.split('\n')
        for line in lines:
            if y < margin + 1*inch: # New Page
                c.showPage()
                # Re-add border on new page
                c.setStrokeColor(HexColor('#2980B9'))
                c.rect(margin, margin, width-2*margin, height-2*margin)
                y = height - margin - 0.5*inch
                c.setFont("Helvetica", 10)
                c.setFillColor(HexColor('#34495E'))

            line = line.strip()
            if not line: continue
            
            # Simple bullet handling
            indent = margin + 0.3*inch
            if line.startswith(('-', '•')):
                c.drawString(indent, y, "•")
                indent += 0.15*inch
                line = line[1:].strip()
            
            # Wrap text manually
            words = line.split()
            current_line = []
            for word in words:
                test_line = " ".join(current_line + [word])
                if c.stringWidth(test_line, "Helvetica", 10) < (width - margin*2 - 0.6*inch):
                    current_line.append(word)
                else:
                    c.drawString(indent, y, " ".join(current_line))
                    y -= 0.22*inch
                    current_line = [word]
                    if y < margin + 0.5*inch: break
            
            if current_line:
                c.drawString(indent, y, " ".join(current_line))
                y -= 0.3*inch # Increased space between matter/paragraphs

    add_section("PROFILE SUMMARY", data["Profile Summary"])
    add_section("PROFESSIONAL EXPERIENCE", data["Professional Experience"])
    add_section("PROJECT EXPERIENCE", data["Project Experience"])
    add_section("TECHNICAL SKILLS", data["Technical Skills"])
    add_section("SOFT SKILLS", data["Soft Skills"])

    # 5. Footer (Left aligned)
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(black)
    c.drawString(margin + 0.2*inch, margin + 0.4*inch, "Private and Confidential Document – Intended for authorized organizations only.")
    c.setFont("Helvetica", 8)
    c.drawString(margin + 0.2*inch, margin + 0.25*inch, "info@krify.com")

    c.save()
    return path

@app.route("/")
def index(): return render_template("hr_converter.html")

@app.route("/convert", methods=["POST"])
def convert_route():
    file = request.files.get("candidate_resume")
    if not file: return redirect("/")
    
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    in_path = os.path.join(TEMP_DIR, f"in_{ts}_{file.filename}")
    file.save(in_path)
    
    cleanup_list = [in_path]
    try:
        txt = extract_text(in_path)
        data = get_ai_data(txt)
        if not data: return redirect("/")
        
        safe_name = re.sub(r'[^a-zA-Z0-9]', '_', data["Full Name"])[:30] or "Candidate"
        out_path = os.path.join(TEMP_DIR, f"Resume_{safe_name}_{ts}.pdf")
        final_path = create_resume_pdf(data, out_path)
        
        cleanup_list.append(final_path)
        
        @after_this_request
        def cleanup(resp):
            for f in cleanup_list:
                if os.path.exists(f): 
                    try: os.remove(f)
                    except: pass
            return resp
            
        return send_file(final_path, as_attachment=True, download_name=os.path.basename(final_path))
    except Exception as e:
        print(f"Route Error: {e}")
        return redirect("/")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5003))
    app.run(host="0.0.0.0", port=port)
