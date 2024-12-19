/*
 * Mimic clipboard behavior for Outlook Mobile by rendering and inserting the signature.
 */

Office.onReady();

async function loadSignatureFromFile() {
    const filePath = "https://siggy.wearelegence.com/users/corey.gashlin@wearelegence.com.html";
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

function renderHtmlForClipboard(html) {
    // Render the HTML in a hidden container
    const container = document.createElement("div");
    container.style.visibility = "hidden";
    container.innerHTML = html;

    // Ensure all images are embedded as base64 (if not already)
    const images = container.querySelectorAll("img");
    images.forEach((img) => {
        if (!img.src.startsWith("data:image")) {
            img.src = convertImageToBase64(img.src);
        }
    });

    document.body.appendChild(container);

    // Extract the fully rendered innerHTML (clipboard-like content)
    const processedContent = container.innerHTML;
    document.body.removeChild(container);

    return processedContent;
}

async function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();

    // Load and process the signature file
    const rawHtmlSignature = await loadSignatureFromFile();
    if (!rawHtmlSignature) {
        console.error("Failed to load the signature.");
        event.completed();
        return;
    }

    // Render HTML for clipboard-like insertion
    const renderedHtmlSignature = renderHtmlForClipboard(rawHtmlSignature);

    // Insert the rendered HTML
    item.body.setSignatureAsync(renderedHtmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error(result.error.message);
        }
        event.completed();
    });

    // Add a notification to confirm
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Signature added successfully",
        icon: "none",
        persistent: false
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("signatureNotification", notification);
}

// Helper function to convert image URLs to base64
function convertImageToBase64(imageUrl) {
    // Fetch the image and convert to base64 (requires server-side or CORS support)
    return ""; // Placeholder - implement as needed or ensure all images are already base64
}
