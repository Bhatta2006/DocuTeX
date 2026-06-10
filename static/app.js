document.addEventListener("DOMContentLoaded", () => {
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const uploadPrompt = document.getElementById("upload-prompt");
    const compileBtn = document.getElementById("compile-btn");
    const statusContainer = document.getElementById("status-container");
    const logSection = document.getElementById("log-section");
    const terminalLogs = document.getElementById("terminal-logs");
    const downloadBtn = document.getElementById("download-btn");

    let currentFile = null;

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
        compileBtn.disabled = false;
        hideStatus();
        logSection.hidden = true;
        downloadBtn.hidden = true;
    }

    function resetFile() {
        currentFile = null;
        fileInput.value = "";
        uploadPrompt.textContent = "Select .tex file";
        compileBtn.disabled = true;
    }

    compileBtn.addEventListener("click", async () => {
        if (!currentFile) return;

        compileBtn.disabled = true;
        downloadBtn.hidden = true;
        logSection.hidden = false;
        terminalLogs.innerHTML = "";
        showStatus("Status: Compiling...");

        appendLogLine("Initializing compiler subprocess...", "system-msg");
        appendLogLine(`Processing ${currentFile.name.toUpperCase()}...`, "system-msg");

        const formData = new FormData();
        formData.append("file", currentFile);

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
});
