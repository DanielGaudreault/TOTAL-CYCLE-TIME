document.getElementById("uploadForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const formData = new FormData();
    formData.append("excel", document.getElementById("excel").files[0]);

    const pdfFiles = document.getElementById("pdfs").files;
    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append("pdfs", pdfFiles[i]);
    }

    try {
        const response = await fetch("http://127.0.0.1:5000/upload", {
            method: "POST",
            body: formData,
        });

        const result = await response.json();
        document.getElementById("result").innerText = result.message;
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerText = "An error occurred.";
    }
});
