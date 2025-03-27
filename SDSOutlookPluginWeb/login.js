
// Basic encoding for light token obfuscation
function encrypt(str) {
    return btoa(unescape(encodeURIComponent(str)));
}

function decrypt(str) {
    try {
        return decodeURIComponent(escape(atob(str)));
    } catch {
        return null;
    }
}

async function doLogin() {
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    const payload = {
        userName: username,
        password: password,
        loginRsponse: {
            passwordRestrinctionVerifiedStatus: {
                nbRepete: 0,
                isVerified: true,
                message: "Password respect requirements"
            }
        }
    };

    try {
        const response = await fetch('https://authentication.tlc-com.com/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const result = await response.json();

        const accessToken = result?.token?.access_token;
        const refreshToken = result?.token?.refresh_token;

        if (accessToken && refreshToken) {
            const settings = Office.context?.roamingSettings;
            if (settings) {
                settings.set("access_token", encrypt(accessToken));
                settings.set("refresh_token", encrypt(refreshToken));
                settings.set("erp_username", result?.userName || username);
                settings.saveAsync(() => {
                    showLoggedInUI(result?.userName || username);
                });
            }
        } else {
            document.getElementById('result').innerText = "❌ Login Failed.";
        }
    } catch (error) {
        document.getElementById('result').innerText = "🚨 Error: " + error.message;
    }
}

function logout() {
    const settings = Office.context?.roamingSettings;
    if (settings) {
        settings.remove("access_token");
        settings.remove("refresh_token");
        settings.remove("erp_username");
        settings.saveAsync(() => {
            showLoginUI();
        });
    }
}

function showLoginUI() {
    document.getElementById('login-form').style.display = 'block';
    document.getElementById('logged-in-section').style.display = 'none';
    document.getElementById('title').innerText = "Login to ERP";
    document.getElementById('result').innerText = "";
}

function showLoggedInUI(username) {
    document.getElementById('login-form').style.display = 'none';
    document.getElementById('logged-in-section').style.display = 'block';
    document.getElementById('welcome-msg').innerText = `✅ You are logged in as: ${username}`;
    document.getElementById('title').innerText = "Welcome!";
    document.getElementById('result').innerText = "";
}

// Utility function to generate a random file guid
function generateRandomFileGuid() {
    // Generates a random 6-digit number as a string
    return Math.floor(100000 + Math.random() * 900000).toString();
}

// Function to upload the file
async function uploadFile(file, fileNumber) {
    const reader = new FileReader();
    reader.onload = async function (event) {
        // Remove the data URL prefix from the result
        const base64String = event.target.result.split(',')[1];
        const originalFileName = file.name;
        const dotIndex = originalFileName.lastIndexOf('.');
        const namePart = dotIndex !== -1 ? originalFileName.substring(0, dotIndex) : originalFileName;
        const extension = dotIndex !== -1 ? originalFileName.substring(dotIndex + 1).toLowerCase() : '';
        const fileSize = file.size;
        const fileGuid = generateRandomFileGuid(extension);

        const payload = {
            interfaceName: "OpFiles",
            code: fileNumber,  // file number from the textbox
            secondCode: null,
            typeCode: 0,
            files: [
                {
                    fileName: originalFileName,  // original file name
                    fileCode: null,
                    fileType: 0,
                    fileCategory: 0,
                    fileExtension: extension,   // file extension extracted from name
                    fileSize: fileSize,
                    fileGuid: fileGuid,         // random filename generated
                    fileBase64: base64String,
                    recognized: false
                }
            ],
            documentId: null
        };

        try {
            const response = await fetch('https://edm.tlc-com.com/EDMComponent/UploadFiles', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await response.json();
            console.log(result);
            if (result && result.interfaceName === "OpFiles") {
                document.getElementById('result').innerText = "✅ OK";
            } else {
                document.getElementById('result').innerText = "❌ " + (result.message || "Upload failed.");
            }
        } catch (error) {
            document.getElementById('result').innerText = "🚨 Error: " + error.message;
        }
    };
    reader.readAsDataURL(file);
}

// Set up drag and drop functionality for file upload
function setupDragAndDrop() {
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('file-input');

    dropArea.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            const file = e.target.files[0];
            const fileNumber = document.getElementById('file-number').value.trim();
            if (!fileNumber) {
                document.getElementById('result').innerText = "⚠️ Please enter a file number.";
                return;
            }
            uploadFile(file, fileNumber);
        }
    });

    dropArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropArea.style.borderColor = "#000";
    });

    dropArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropArea.style.borderColor = "#ccc";
    });

    dropArea.addEventListener('drop', (e) => {
        e.preventDefault();
        dropArea.style.borderColor = "#ccc";
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            const fileNumber = document.getElementById('file-number').value.trim();
            if (!fileNumber) {
                document.getElementById('result').innerText = "⚠️ Please enter a file number.";
                return;
            }
            uploadFile(file, fileNumber);
        }
    });
}

Office.onReady(() => {
    checkLoginStatus();
    setupDragAndDrop();
});

function checkLoginStatus() {
    const settings = Office.context?.roamingSettings;
    if (!settings) {
        showLoginUI();
        return;
    }

    const encryptedToken = settings.get("access_token");
    const username = settings.get("erp_username");

    if (encryptedToken && username) {
        const token = decrypt(encryptedToken);
        if (token) {
            showLoggedInUI(username);
        } else {
            showLoginUI();
        }
    } else {
        showLoginUI();
    }
}
