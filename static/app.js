document.addEventListener("DOMContentLoaded", () => {
    const tabFile = document.getElementById("tab-file");
    const tabCode = document.getElementById("tab-code");
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const uploadPrompt = document.getElementById("upload-prompt");
    const codeAreaContainer = document.getElementById("code-area-container");
    const latexCode = document.getElementById("latex-code");
    const compileBtn = document.getElementById("compile-btn");
    const statusContainer = document.getElementById("status-container");
    const logSection = document.getElementById("log-section");
    const terminalLogs = document.getElementById("terminal-logs");
    const downloadBtn = document.getElementById("download-btn");

    let currentFile = null;
    let activeTab = "file"; // "file" or "code"

    // Tab Navigation Handlers
    tabFile.addEventListener("click", () => {
        if (activeTab === "file") return;
        activeTab = "file";
        tabFile.classList.add("active");
        tabCode.classList.remove("active");
        dropZone.hidden = false;
        codeAreaContainer.hidden = true;
        updateCompileButtonState();
    });

    tabCode.addEventListener("click", () => {
        if (activeTab === "code") return;
        activeTab = "code";
        tabCode.classList.add("active");
        tabFile.classList.remove("active");
        dropZone.hidden = true;
        codeAreaContainer.hidden = false;
        updateCompileButtonState();
    });

    // Helper to toggle Compile Button disabled state
    function updateCompileButtonState() {
        if (activeTab === "file") {
            compileBtn.disabled = !currentFile;
        } else {
            compileBtn.disabled = !latexCode.value.trim();
        }
    }

    latexCode.addEventListener("input", () => {
        updateCompileButtonState();
    });

    dropZone.addEventListener("click", () => {
        fileInput.click();
    });

    fileInput.addEventListener("change", (e) => {
        handleFiles(e.target.files);
    });

    ["dragenter", "dragover"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add("dragover");
        }, false);
    });

    ["dragleave", "drop"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove("dragover");
        }, false);
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        handleFiles(e.dataTransfer.files);
    });

    function handleFiles(files) {
        if (files.length === 0) return;
        const file = files[0];
        
        if (!file.name.endsWith(".tex")) {
            showStatus("Error: Only .tex files allowed");
            resetFile();
            return;
        }

        currentFile = file;
        uploadPrompt.textContent = file.name;
        updateCompileButtonState();
        hideStatus();
        logSection.hidden = true;
        downloadBtn.hidden = true;
    }

    function resetFile() {
        currentFile = null;
        fileInput.value = "";
        uploadPrompt.textContent = "Select .tex file";
        updateCompileButtonState();
    }

    compileBtn.addEventListener("click", async () => {
        let fileToUpload = null;
        let filenameForLogs = "";

        if (activeTab === "file") {
            if (!currentFile) return;
            fileToUpload = currentFile;
            filenameForLogs = currentFile.name.toUpperCase();
        } else {
            const codeText = latexCode.value.trim();
            if (!codeText) return;
            
            // Convert pasted text into a virtual file
            const blob = new Blob([codeText], { type: "application/x-latex" });
            fileToUpload = new File([blob], "pasted.tex", { type: "application/x-latex" });
            filenameForLogs = "PASTED LATEX CODE";
        }

        compileBtn.disabled = true;
        downloadBtn.hidden = true;
        logSection.hidden = false;
        terminalLogs.innerHTML = "";
        showStatus("Status: Compiling...");

        appendLogLine("Initializing compiler subprocess...", "system-msg");
        appendLogLine(`Processing ${filenameForLogs}...`, "system-msg");

        const formData = new FormData();
        formData.append("file", fileToUpload);

        try {
            const response = await fetch("/convert", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Server error.");
            }

            const data = await response.json();
            
            // Print logs to terminal
            if (data.logs && data.logs.length > 0) {
                for (let line of data.logs) {
                    const type = line.toLowerCase().includes("warning") || line.toLowerCase().includes("error") ? "stderr" : "stdout";
                    appendLogLine(line, type);
                }
            }

            if (data.success) {
                showStatus("Status: Success");
                downloadBtn.href = data.download_url;
                downloadBtn.hidden = false;
                appendLogLine("Compilation completed.", "system-msg");
            } else {
                showStatus(`Error: Compilation failed.`);
                appendLogLine(`Fatal exception: ${data.error || "Execution error."}`, "stderr");
            }

        } catch (error) {
            showStatus(`Error: ${error.message}`);
            appendLogLine(`System error: ${error.message}`, "stderr");
        } finally {
            compileBtn.disabled = false;
        }
    });

    function showStatus(msg) {
        statusContainer.textContent = msg;
        statusContainer.hidden = false;
    }

    function hideStatus() {
        statusContainer.hidden = true;
        statusContainer.textContent = "";
    }

    function appendLogLine(text, type = "stdout") {
        const line = document.createElement("div");
        line.className = `log-line ${type}`;
        line.textContent = `> ${text}`;
        terminalLogs.appendChild(line);
        terminalLogs.scrollTop = terminalLogs.scrollHeight;
    }

    // GitHub Link & Auto-Star Interaction Handler
    const githubLink = document.getElementById("github-link");
    const toast = document.getElementById("toast");

    githubLink.addEventListener("click", (e) => {
        e.preventDefault();
        
        // Display redirect & star reminder toast
        toast.innerHTML = "Redirecting to GitHub...<br>Please press the <strong>STAR ⭐</strong> button on the top right to support the project!";
        toast.hidden = false;

        // Open repository in a new window after 1.8 seconds delay
        setTimeout(() => {
            window.open(githubLink.href, "_blank");
            
            // Fade out/hide toast after another few seconds
            setTimeout(() => {
                toast.hidden = true;
            }, 3000);
        }, 1800);
    });
});
