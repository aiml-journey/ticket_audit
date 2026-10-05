document.getElementById("auditForm").addEventListener(
    "submit",
    async (e) => {

        e.preventDefault();

        const btn =
            document.getElementById("submitBtn");

        const misFile =
            document.getElementById("misFile").files[0];

        const pdfFiles =
            document.getElementById("pdfFiles").files;


        if (!misFile) {
            alert("Please select the MIS Excel file.");
            return;
        }

        if (pdfFiles.length === 0) {
            alert("Please select at least one PDF ticket.");
            return;
        }


        btn.textContent =
            "Auditing tickets...";

        btn.disabled = true;


        const formData =
            new FormData();

        formData.append(
            "mis_file",
            misFile
        );


        for (
            const file of pdfFiles
        ) {

            formData.append(
                "pdf_files",
                file
            );

        }


        try {

            const response =
                await fetch(
                    "/api/audit",
                    {
                        method: "POST",
                        body: formData
                    }
                );


            if (!response.ok) {

                let message =
                    "Audit failed.";

                try {

                    const error =
                        await response.json();

                    message =
                        error.details ||
                        error.error ||
                        message;

                } catch (_) {}

                throw new Error(
                    message
                );
            }


            const blob =
                await response.blob();


            if (
                !blob ||
                blob.size === 0
            ) {

                throw new Error(
                    "The server returned an empty Excel file."
                );

            }


            const url =
                window.URL.createObjectURL(
                    blob
                );


            const link =
                document.createElement("a");

            link.href = url;

            link.download =
                "Audited_" +
                misFile.name;


            document.body.appendChild(
                link
            );

            link.click();

            link.remove();


            setTimeout(() => {

                window.URL.revokeObjectURL(
                    url
                );

            }, 1000);


            btn.textContent =
                "Audit Complete ✓";


        } catch (error) {

            console.error(
                error
            );

            alert(
                "Audit failed:\n\n" +
                error.message
            );

            btn.textContent =
                "Run Audit & Download Updated Excel Sheet";


        } finally {

            btn.disabled = false;

        }

    }
);
