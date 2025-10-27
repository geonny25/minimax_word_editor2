// Word Editor - Main JavaScript File
class WordEditor {
  constructor() {
    this.editor = document.getElementById("editor")
    this.documentTitle = document.getElementById("documentTitle")
    this.isDarkMode = localStorage.getItem("darkMode") === "true"
    this.history = []
    this.historyStep = -1
    this.autoSaveInterval = null
    this.lastSavedContent = ""

    this.init()
  }

  init() {
    this.setupEventListeners()
    this.loadDocument()
    this.startAutoSave()
    this.updateStats()

    if (this.isDarkMode) {
      this.toggleDarkMode()
    }
  }

  setupEventListeners() {
    // File Operations
    document.getElementById("newBtn").addEventListener("click", () => this.newDocument())
    document.getElementById("openBtn").addEventListener("click", () => this.openFile())
    document.getElementById("saveBtn").addEventListener("click", () => this.saveDocument())
    document.getElementById("printBtn").addEventListener("click", () => this.printDocument())

    // Edit Operations
    document.getElementById("undoBtn").addEventListener("click", () => this.undo())
    document.getElementById("redoBtn").addEventListener("click", () => this.redo())
    document.getElementById("cutBtn").addEventListener("click", () => document.execCommand("cut"))
    document.getElementById("copyBtn").addEventListener("click", () => document.execCommand("copy"))
    document.getElementById("pasteBtn").addEventListener("click", () => document.execCommand("paste"))

    // Find & Replace
    document.getElementById("findBtn").addEventListener("click", () => this.openModal("findReplaceModal"))
    document.getElementById("replaceBtn").addEventListener("click", () => this.openModal("findReplaceModal"))

    // Font Formatting
    document.getElementById("fontFamily").addEventListener("change", (e) => {
      document.execCommand("fontName", false, e.target.value)
      this.editor.focus()
    })
    document.getElementById("fontSize").addEventListener("change", (e) => {
      const selection = window.getSelection()
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0)
        const span = document.createElement("span")
        span.style.fontSize = e.target.value
        range.surroundContents(span)
      }
      this.editor.focus()
    })

    // Text Formatting
    document.getElementById("boldBtn").addEventListener("click", () => {
      document.execCommand("bold")
      this.editor.focus()
    })
    document.getElementById("italicBtn").addEventListener("click", () => {
      document.execCommand("italic")
      this.editor.focus()
    })
    document.getElementById("underlineBtn").addEventListener("click", () => {
      document.execCommand("underline")
      this.editor.focus()
    })
    document.getElementById("strikeBtn").addEventListener("click", () => {
      document.execCommand("strikethrough")
      this.editor.focus()
    })
    document.getElementById("textColor").addEventListener("change", (e) => {
      document.execCommand("foreColor", false, e.target.value)
      this.editor.focus()
    })
    document.getElementById("highlightColor").addEventListener("change", (e) => {
      document.execCommand("backColor", false, e.target.value)
      this.editor.focus()
    })

    // Alignment
    document.getElementById("alignLeftBtn").addEventListener("click", () => {
      document.execCommand("justifyLeft")
      this.editor.focus()
    })
    document.getElementById("alignCenterBtn").addEventListener("click", () => {
      document.execCommand("justifyCenter")
      this.editor.focus()
    })
    document.getElementById("alignRightBtn").addEventListener("click", () => {
      document.execCommand("justifyRight")
      this.editor.focus()
    })
    document.getElementById("alignJustifyBtn").addEventListener("click", () => {
      document.execCommand("justifyFull")
      this.editor.focus()
    })

    // Lists
    document.getElementById("bulletBtn").addEventListener("click", () => {
      document.execCommand("insertUnorderedList")
      this.editor.focus()
    })
    document.getElementById("numberBtn").addEventListener("click", () => {
      document.execCommand("insertOrderedList")
      this.editor.focus()
    })
    document.getElementById("increaseIndentBtn").addEventListener("click", () => {
      document.execCommand("indent")
      this.editor.focus()
    })
    document.getElementById("decreaseIndentBtn").addEventListener("click", () => {
      document.execCommand("outdent")
      this.editor.focus()
    })

    // Insert
    document.getElementById("insertImageBtn").addEventListener("click", () => this.openModal("imageModal"))
    document.getElementById("insertTableBtn").addEventListener("click", () => this.openModal("tableModal"))
    document.getElementById("insertLinkBtn").addEventListener("click", () => this.openModal("linkModal"))
    document.getElementById("insertLineBtn").addEventListener("click", () => this.insertHorizontalLine())

    // Advanced
    document.getElementById("templateBtn").addEventListener("click", () => this.openModal("templatesModal"))
    document.getElementById("statsBtn").addEventListener("click", () => this.showStats())
    document.getElementById("exportBtn").addEventListener("click", () => this.openModal("exportModal"))

    // Dark Mode
    document.getElementById("darkModeToggle").addEventListener("click", () => this.toggleDarkMode())

    // Help
    document.getElementById("helpBtn").addEventListener("click", () => this.openModal("helpModal"))

    // Editor Events
    this.editor.addEventListener("input", () => {
      this.saveHistory()
      this.updateStats()
      this.updateAutoSaveStatus()
    })

    // Modal Close Buttons
    document.querySelectorAll(".modal-close, .modal-close-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const modal = e.target.closest(".modal")
        if (modal) this.closeModal(modal.id)
      })
    })

    // Find & Replace
    document.getElementById("findNextBtn").addEventListener("click", () => this.findNext())
    document.getElementById("replaceBtn2").addEventListener("click", () => this.replace())
    document.getElementById("replaceAllBtn").addEventListener("click", () => this.replaceAll())

    // Insert Table
    document.getElementById("insertTableConfirmBtn").addEventListener("click", () => this.insertTable())

    // Insert Link
    document.getElementById("insertLinkConfirmBtn").addEventListener("click", () => this.insertLink())

    // Insert Image
    document.getElementById("insertImageConfirmBtn").addEventListener("click", () => this.insertImage())

    // Templates
    document.querySelectorAll(".template-card").forEach((card) => {
      card.addEventListener("click", () => this.loadTemplate(card.dataset.template))
    })

    // Export
    document.getElementById("exportDocxBtn").addEventListener("click", () => this.exportAsDocx())
    document.getElementById("exportTxtBtn").addEventListener("click", () => this.exportAsText())
    document.getElementById("exportHtmlBtn").addEventListener("click", () => this.exportAsHTML())
    document.getElementById("exportPdfBtn").addEventListener("click", () => this.exportAsPDF())

    // Document Title
    this.documentTitle.addEventListener("change", () => this.saveDocument())

    // Keyboard Shortcuts
    document.addEventListener("keydown", (e) => this.handleKeyboardShortcuts(e))

    // Close modals on escape
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        document.querySelectorAll(".modal.active").forEach((modal) => {
          this.closeModal(modal.id)
        })
      }
    })
  }

  // File Operations
  newDocument() {
    if (
      this.editor.innerHTML.trim() !== "<p>Start typing your document here...</p>" &&
      !confirm("Are you sure? Any unsaved changes will be lost.")
    ) {
      return
    }
    this.editor.innerHTML = "<p>Start typing your document here...</p>"
    this.documentTitle.value = "Untitled Document"
    this.history = []
    this.historyStep = -1
    this.updateStats()
    localStorage.removeItem("editorContent")
    localStorage.removeItem("editorTitle")
  }

  openFile() {
    const input = document.createElement("input")
    input.type = "file"
    input.accept = ".txt,.html,.docx"
    input.addEventListener("change", (e) => {
      const file = e.target.files[0]
      if (!file) return

      const reader = new FileReader()
      reader.onload = (event) => {
        const content = event.target.result
        if (file.name.endsWith(".html")) {
          this.editor.innerHTML = content
        } else {
          this.editor.textContent = content
        }
        this.documentTitle.value = file.name.replace(/\.[^/.]+$/, "")
        this.history = []
        this.historyStep = -1
        this.updateStats()
      }
      reader.readAsText(file)
    })
    input.click()
  }

  saveDocument() {
    const content = this.editor.innerHTML
    const title = this.documentTitle.value || "Untitled Document"

    localStorage.setItem("editorContent", content)
    localStorage.setItem("editorTitle", title)
    this.lastSavedContent = content

    this.updateAutoSaveStatus("Saved")
    setTimeout(() => this.updateAutoSaveStatus("Auto-saved"), 2000)
  }

  loadDocument() {
    const savedContent = localStorage.getItem("editorContent")
    const savedTitle = localStorage.getItem("editorTitle")

    if (savedContent) {
      this.editor.innerHTML = savedContent
      this.lastSavedContent = savedContent
    }
    if (savedTitle) {
      this.documentTitle.value = savedTitle
    }
  }

  printDocument() {
    const printWindow = window.open("", "", "height=600,width=800")
    printWindow.document.write("<html><head><title>" + this.documentTitle.value + "</title>")
    printWindow.document.write("<style>body { font-family: Arial, sans-serif; line-height: 1.6; margin: 2cm; }")
    printWindow.document.write("h1, h2, h3 { margin-top: 1em; } table { border-collapse: collapse; width: 100%; }")
    printWindow.document.write(
      "td, th { border: 1px solid #ddd; padding: 8px; text-align: left; }</style></head><body>",
    )
    printWindow.document.write(this.editor.innerHTML)
    printWindow.document.write("</body></html>")
    printWindow.document.close()
    printWindow.print()
  }

  // History Management
  saveHistory() {
    this.historyStep++
    if (this.historyStep < this.history.length) {
      this.history.length = this.historyStep
    }
    this.history.push(this.editor.innerHTML)
  }

  undo() {
    if (this.historyStep > 0) {
      this.historyStep--
      this.editor.innerHTML = this.history[this.historyStep]
      this.updateStats()
    }
  }

  redo() {
    if (this.historyStep < this.history.length - 1) {
      this.historyStep++
      this.editor.innerHTML = this.history[this.historyStep]
      this.updateStats()
    }
  }

  // Find & Replace
  findNext() {
    const findText = document.getElementById("findInput").value
    const caseSensitive = document.getElementById("caseSensitive").checked

    if (!findText) return

    const content = this.editor.textContent
    const searchText = caseSensitive ? findText : findText.toLowerCase()
    const contentToSearch = caseSensitive ? content : content.toLowerCase()

    const index = contentToSearch.indexOf(searchText)
    if (index !== -1) {
      const range = document.createRange()
      const sel = window.getSelection()

      let charCount = 0
      const nodeStack = [this.editor]
      let node,
        foundStart = false,
        stop = false

      while (!stop && (node = nodeStack.pop())) {
        if (node.nodeType === 3) {
          const nextCharCount = charCount + node.length
          if (!foundStart && index < nextCharCount) {
            foundStart = true
            range.setStart(node, index - charCount)
          }
          if (foundStart && index + findText.length <= nextCharCount) {
            range.setEnd(node, index + findText.length - charCount)
            stop = true
          }
          charCount = nextCharCount
        } else {
          let i = node.childNodes.length
          while (i--) {
            nodeStack.push(node.childNodes[i])
          }
        }
      }

      sel.removeAllRanges()
      sel.addRange(range)
    }
  }

  replace() {
    const findText = document.getElementById("findInput").value
    const replaceText = document.getElementById("replaceInput").value

    if (!findText) return

    const sel = window.getSelection()
    if (sel.toString() === findText) {
      document.execCommand("insertText", false, replaceText)
      this.findNext()
    }
  }

  replaceAll() {
    const findText = document.getElementById("findInput").value
    const replaceText = document.getElementById("replaceInput").value

    if (!findText) return

    let content = this.editor.innerHTML
    const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g")
    content = content.replace(regex, replaceText)
    this.editor.innerHTML = content
    this.saveHistory()
    this.updateStats()
  }

  // Insert Elements
  insertTable() {
    const rows = Number.parseInt(document.getElementById("tableRows").value) || 3
    const cols = Number.parseInt(document.getElementById("tableCols").value) || 3

    let table = "<table><tbody>"
    for (let i = 0; i < rows; i++) {
      table += "<tr>"
      for (let j = 0; j < cols; j++) {
        table += "<td>Cell</td>"
      }
      table += "</tr>"
    }
    table += "</tbody></table><p></p>"

    const selection = window.getSelection()
    if (selection.rangeCount > 0) {
      const range = selection.getRangeAt(0)
      const fragment = document.createRange().createContextualFragment(table)
      range.insertNode(fragment)
    }

    this.closeModal("tableModal")
    this.saveHistory()
    this.editor.focus()
  }

  insertLink() {
    const linkText = document.getElementById("linkText").value || "Link"
    const linkUrl = document.getElementById("linkUrl").value || "#"

    const link = `<a href="${linkUrl}" target="_blank">${linkText}</a>`
    document.execCommand("insertHTML", false, link)
    this.closeModal("linkModal")
    this.saveHistory()
    this.editor.focus()
  }

  insertImage() {
    const imageFile = document.getElementById("imageFile").files[0]
    const imageAlt = document.getElementById("imageAlt").value || "Image"
    const imageWidth = document.getElementById("imageWidth").value || "300"

    if (!imageFile) {
      alert("Please select an image file")
      return
    }

    const reader = new FileReader()
    reader.onload = (e) => {
      const imageUrl = e.target.result
      const img = `<img src="${imageUrl}" alt="${imageAlt}" style="max-width: ${imageWidth}px; height: auto;">`
      document.execCommand("insertHTML", false, img)
      this.saveHistory()
    }
    reader.readAsDataURL(imageFile)

    this.closeModal("imageModal")
    this.editor.focus()
  }

  insertHorizontalLine() {
    document.execCommand("insertHTML", false, "<hr>")
    this.saveHistory()
    this.editor.focus()
  }

  // Templates
  loadTemplate(templateType) {
    const templates = {
      resume: `<h1>Your Name</h1>
<p>Email: your.email@example.com | Phone: (123) 456-7890</p>
<h2>Professional Summary</h2>
<p>Brief overview of your professional background and goals.</p>
<h2>Experience</h2>
<h3>Job Title - Company Name</h3>
<p>Duration: Month Year - Present</p>
<ul><li>Achievement or responsibility</li><li>Achievement or responsibility</li></ul>
<h2>Education</h2>
<h3>Degree - University Name</h3>
<p>Graduation: Month Year</p>
<h2>Skills</h2>
<ul><li>Skill 1</li><li>Skill 2</li><li>Skill 3</li></ul>`,

      letter: `<p>Your Address<br>City, State ZIP<br>Date</p>
<p>Recipient Name<br>Company Name<br>Address<br>City, State ZIP</p>
<p>Dear [Recipient Name],</p>
<p>Opening paragraph introducing the purpose of the letter.</p>
<p>Body paragraph with main content and details.</p>
<p>Closing paragraph with call to action.</p>
<p>Sincerely,<br><br>Your Name</p>`,

      report: `<h1>Report Title</h1>
<p><strong>Date:</strong> [Date]<br><strong>Author:</strong> [Your Name]</p>
<h2>Executive Summary</h2>
<p>Brief overview of the report findings and recommendations.</p>
<h2>Introduction</h2>
<p>Background and context for the report.</p>
<h2>Findings</h2>
<h3>Finding 1</h3>
<p>Details about the first finding.</p>
<h3>Finding 2</h3>
<p>Details about the second finding.</p>
<h2>Recommendations</h2>
<ul><li>Recommendation 1</li><li>Recommendation 2</li></ul>
<h2>Conclusion</h2>
<p>Summary and final thoughts.</p>`,
    }

    this.editor.innerHTML = templates[templateType] || ""
    this.closeModal("templatesModal")
    this.saveHistory()
    this.updateStats()
  }

  // Statistics
  updateStats() {
    const text = this.editor.innerText
    const words = text
      .trim()
      .split(/\s+/)
      .filter((w) => w.length > 0).length
    const chars = text.length
    const charsNoSpaces = text.replace(/\s/g, "").length
    const lines = text.split("\n").length
    const paragraphs = this.editor.querySelectorAll("p").length
    const readingTime = Math.ceil(words / 200)

    document.getElementById("wordCount").textContent = `Words: ${words}`
    document.getElementById("charCount").textContent = `Characters: ${chars}`
    document.getElementById("lineCount").textContent = `Lines: ${lines}`
  }

  showStats() {
    const text = this.editor.innerText
    const words = text
      .trim()
      .split(/\s+/)
      .filter((w) => w.length > 0).length
    const charsWithSpaces = text.length
    const charsWithoutSpaces = text.replace(/\s/g, "").length
    const paragraphs = this.editor.querySelectorAll("p").length
    const lines = text.split("\n").length
    const readingTime = Math.ceil(words / 200)

    document.getElementById("statsWords").textContent = words
    document.getElementById("statsCharsWithSpaces").textContent = charsWithSpaces
    document.getElementById("statsCharsWithoutSpaces").textContent = charsWithoutSpaces
    document.getElementById("statsParagraphs").textContent = paragraphs
    document.getElementById("statsLines").textContent = lines
    document.getElementById("statsReadingTime").textContent = `${readingTime} min`

    this.openModal("statsModal")
  }

  // Export
  exportAsDocx() {
    const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, AlignmentType } = window.docx

    const htmlContent = this.editor.innerHTML
    const tempDiv = document.createElement("div")
    tempDiv.innerHTML = htmlContent

    const paragraphs = []
    const processNode = (node) => {
      if (node.nodeType === 3) {
        const text = node.textContent.trim()
        if (text) {
          paragraphs.push(new Paragraph({ text }))
        }
      } else if (node.nodeType === 1) {
        const tagName = node.tagName.toLowerCase()
        if (tagName === "p") {
          paragraphs.push(new Paragraph({ text: node.innerText || "" }))
        } else if (tagName === "h1") {
          paragraphs.push(new Paragraph({ text: node.innerText || "", size: 32 }))
        } else if (tagName === "h2") {
          paragraphs.push(new Paragraph({ text: node.innerText || "", size: 28 }))
        } else if (tagName === "h3") {
          paragraphs.push(new Paragraph({ text: node.innerText || "", size: 24 }))
        } else if (tagName === "ul" || tagName === "ol") {
          node.querySelectorAll("li").forEach((li) => {
            paragraphs.push(new Paragraph({ text: li.innerText || "", bullet: { level: 0 } }))
          })
        } else if (tagName === "table") {
          const rows = []
          node.querySelectorAll("tr").forEach((tr) => {
            const cells = []
            tr.querySelectorAll("td, th").forEach((td) => {
              cells.push(new TableCell({ children: [new Paragraph(td.innerText || "")] }))
            })
            rows.push(new TableRow({ children: cells }))
          })
          paragraphs.push(new Table({ rows }))
        } else {
          node.childNodes.forEach(processNode)
        }
      }
    }

    tempDiv.childNodes.forEach(processNode)

    const doc = new Document({ sections: [{ children: paragraphs }] })

    Packer.toBlob(doc).then((blob) => {
      const url = URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.href = url
      a.download = this.documentTitle.value + ".docx"
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
    })

    this.closeModal("exportModal")
  }

  exportAsText() {
    const text = this.editor.innerText
    const filename = this.documentTitle.value + ".txt"
    this.downloadFile(text, filename, "text/plain")
    this.closeModal("exportModal")
  }

  exportAsHTML() {
    const html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${this.documentTitle.value}</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 2cm; }
        h1, h2, h3 { margin-top: 1em; }
        table { border-collapse: collapse; width: 100%; }
        td, th { border: 1px solid #ddd; padding: 8px; text-align: left; }
    </style>
</head>
<body>
    ${this.editor.innerHTML}
</body>
</html>`
    const filename = this.documentTitle.value + ".html"
    this.downloadFile(html, filename, "text/html")
    this.closeModal("exportModal")
  }

  exportAsPDF() {
    alert("PDF export requires a library like jsPDF. For now, use your browser's Print to PDF feature (Ctrl+P).")
    this.closeModal("exportModal")
  }

  downloadFile(content, filename, type) {
    const blob = new Blob([content], { type })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  // Dark Mode
  toggleDarkMode() {
    document.body.classList.toggle("dark-mode")
    this.isDarkMode = !this.isDarkMode
    localStorage.setItem("darkMode", this.isDarkMode)
    document.getElementById("darkModeToggle").textContent = this.isDarkMode ? "☀️" : "🌙"
  }

  // Auto Save
  startAutoSave() {
    this.autoSaveInterval = setInterval(() => {
      if (this.editor.innerHTML !== this.lastSavedContent) {
        this.saveDocument()
      }
    }, 30000) // Auto-save every 30 seconds
  }

  updateAutoSaveStatus(status) {
    document.getElementById("autoSaveStatus").textContent = status
  }

  // Modal Management
  openModal(modalId) {
    const modal = document.getElementById(modalId)
    if (modal) {
      modal.classList.add("active")
      // Clear inputs when opening
      if (modalId === "findReplaceModal") {
        document.getElementById("findInput").focus()
      } else if (modalId === "tableModal") {
        document.getElementById("tableRows").focus()
      } else if (modalId === "linkModal") {
        document.getElementById("linkText").focus()
      } else if (modalId === "imageModal") {
        document.getElementById("imageFile").focus()
      }
    }
  }

  closeModal(modalId) {
    const modal = document.getElementById(modalId)
    if (modal) {
      modal.classList.remove("active")
    }
  }

  // Keyboard Shortcuts
  handleKeyboardShortcuts(e) {
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case "n":
          e.preventDefault()
          this.newDocument()
          break
        case "o":
          e.preventDefault()
          this.openFile()
          break
        case "s":
          e.preventDefault()
          this.saveDocument()
          break
        case "p":
          e.preventDefault()
          this.printDocument()
          break
        case "f":
          e.preventDefault()
          this.openModal("findReplaceModal")
          break
        case "h":
          e.preventDefault()
          this.openModal("findReplaceModal")
          break
      }
    }
  }
}

// Initialize the editor when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  new WordEditor()
})
