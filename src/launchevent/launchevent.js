/*
 * Insert a cleaned-up HTML signature to avoid extra whitespace.
 */

Office.onReady();

async function loadSignatureFromFile() {
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    // Encode the email to ensure it's URL-safe
    const encodedEmail = encodeURIComponent(userEmail);
    const filePath = `https://siggy.wearelegence.com/users/${encodedEmail}.html?cb=${new Date().getTime()}`;
    try {
        const response = await fetch(filePath, { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
        }
        return await response.text(); // Raw HTML content
    } catch (error) {
        console.error("Error fetching HTML file:", error);
        return null;
    }
}

function cleanHtmlForWhitespace(html) {
    // Create a hidden container to process the HTML
    const container = document.createElement("div");
    container.style.visibility = "hidden";
    container.innerHTML = html;
    document.body.appendChild(container);

    // Normalize margins and padding for all elements
    const allElements = container.querySelectorAll("*");
    allElements.forEach((el) => {
        el.style.margin = "0";
        el.style.padding = "0";
        el.style.lineHeight = "1.2"; // Adjust as needed for consistent spacing
    });

    // Extract cleaned-up HTML
    const cleanedHtml = container.innerHTML;
    document.body.removeChild(container);

    return cleanedHtml;
}

async function onNewMessageComposeHandler(event) {
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();
    const item = Office.context.mailbox.item;
    if (platform.includes("android") || platform.includes("ios")) {
        // Load and process the signature file
        const rawHtmlSignature = await loadSignatureFromFile();
        if (!rawHtmlSignature) {
            console.error("Failed to load the signature.");
            event.completed();
            return;
        }

        // Clean the HTML to remove extra whitespace
        const cleanedHtmlSignature = cleanHtmlForWhitespace(rawHtmlSignature);

        // Insert the cleaned HTML into the email body
        item.body.setSignatureAsync(cleanedHtmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Error inserting the signature:", result.error.message);
            }
            event.completed();
        });

        // Add a notification to confirm success
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added successfully",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("signatureNotification", notification);
    }

}
