# 🖋️ Lumina: Comparative Manuscript Suite (v12.0)
**A Private, AI-Free Scriptorium for Theologians, Translators, and Literary Scholars.**

Lumina is a specialized, zero-footprint typesetting and research engine designed to bridge the gap between complex comparative study and publisher-ready manuscripts. Whether you are analyzing Quranic translations, comparing drafts of classical literature, or mapping thematic elements across legal texts, Lumina handles the heavy lifting of typography, indexing, and formatting so you can focus entirely on the text.

### 🛡️ Core Philosophy: Zero-Footprint Security
Your research is your intellectual property. Lumina is designed with a strict **Zero-Footprint Architecture**. 
* Documents are generated entirely in your browser's RAM. 
* No manuscript data is ever saved to an external database or cloud server.
* You maintain total control by downloading your project save-states directly to your local machine.

---

### 🚀 Quick Start Guide

**Step 1: Configure Your Project (Sidebar)**
* Enter your Book Title, Author Name, and Copyright Year. 
* **Language Direction:** If your primary source text is Arabic or Hebrew, toggle the **Right-to-Left** formatting switch. Lumina will automatically align and inject native RTL XML into your final Word document to preserve calligraphy.
* **Variant Labels:** Type a comma-separated list of your translations or drafts (e.g., *Yusuf Ali, Pickthall, Sahih* OR *Draft A, Draft B, Final*). These will dynamically populate your dropdown menus.

**Step 2: Build Your Taxonomy & "If-Then" Rules**
* **Keyword Rules**: Standard list of names or key terms.
* **If-Then Patterns**: Dynamic rules like *"If sentence starts with 'Say,' highlight the entire sentence."* Lumina uses smart regex to capture full linguistic contexts.

**Step 3: Draft & Verify**
* Use the **Live Print Preview** with its per-section **Highlight Toggle** to double-check your annotations against the raw source text.
* The **If-Then Engine** will automatically highlight patterns as you type, provided they match your defined rules.

**Step 4: Draft & Verify**
* Use the **Live Print Preview** with its per-section **Highlight Toggle** to double-check AI analysis against the raw source text.

**Step 5: Compile & Publish**
* Click **Compile Master Manuscript**. Lumina will instantly generate an Amazon KDP-compliant `.docx` file featuring a Title Page, Copyright Page, Color Key, Hierarchical Chapters, 60/40 Split Parallel Tables, and an Auto-Generated Concordance Index.

---

### 🧠 Advanced Features

**🕸️ The Knowledge Web (Interactive Concordance)**
Text indexes are great for print, but terrible for discovery. Navigate to the **Knowledge Web** tab to view a live, physics-based node map of your manuscript. Drag thematic nodes to visually explore how your taxonomy connects across different chapters and verses.

**🧘 Zen Mode**
Research requires deep focus. Toggle **Zen Mode** at the top of the sidebar to hide all controls, translations, and sidebars. You will be left with only your Primary Source text and a clean canvas for your Academic Commentary.

**💾 Automated Save States**
Lumina silently saves your project in the background after every keystroke. If your browser crashes or you accidentally close the tab, your work will be exactly where you left it. To move between computers, simply click **Export Full Project** to download a `.json` bundle of your entire manuscript.

---

### 🛠️ Technical Specifications
* **Output:** Microsoft Word `.docx` (XML Injected)
* **Typography Engine:** Mimics the classic 1946 Abdullah Yusuf Ali 60/40 Parallel Layout.
* **Environment:** Python / Streamlit (React frontend).
* **Graphing Engine:** Vis.js Network.
