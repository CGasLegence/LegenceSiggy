/*
* Copy-paste rendered HTML into the email body, retaining layout (tables) and inline styles.
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

function renderHtmlWithStructure(html) {
    const container = document.createElement("div");
    container.style.visibility = "hidden";
    container.innerHTML = html;
    document.body.appendChild(container);

    // Convert the rendered HTML to include inline styles
    const processedHtml = traverseAndApplyStyles(container);
    document.body.removeChild(container);

    return processedHtml;
}

function traverseAndApplyStyles(element) {
    const traverseNode = (node) => {
        if (node.nodeType === Node.TEXT_NODE) {
            return node.textContent; // Keep text content
        }

        if (node.nodeType === Node.ELEMENT_NODE) {
            const tagName = node.tagName.toLowerCase();
            const styles = window.getComputedStyle(node);
            let inlineStyle = "";

            // Extract relevant styles to apply inline
            Array.from(styles).forEach((style) => {
                const value = styles.getPropertyValue(style);
                if (value) {
                    inlineStyle += `${style}:${value};`;
                }
            });

            // Recreate the element with inline styles
            const openingTag = `<${tagName} style="${inlineStyle}">`;
            const closingTag = `</${tagName}>`;
            const children = Array.from(node.childNodes)
                .map(traverseNode)
                .join("");

            return `${openingTag}${children}${closingTag}`;
        }

        return ""; // Ignore other node types
    };

    return traverseNode(element);
}

async function onNewMessageComposeHandler(event) {
    const item = Office.context.mailbox.item;
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();

    const rawHtmlSignature = await loadSignatureFromFile();
    if (!rawHtmlSignature) {
        console.error("Failed to load the signature.");
        event.completed();
        return;
    }

    // Render the HTML and apply inline styles
    const styledHtmlSignature = renderHtmlWithStructure(rawHtmlSignature);

    if (platform.includes("android")) {
        console.log("Running Android-specific logic...");
        item.body.setSignatureAsync(styledHtmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for Android",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("androidNotification", notification);
    } else {
        console.log("Running non-Android logic...");
        item.body.setSignatureAsync(styledHtmlSignature, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
            event.completed();
        });

        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Signature added for non-Android platform",
            icon: "none",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("nonAndroidNotification", notification);
    }
}
