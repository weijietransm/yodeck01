document.addEventListener('DOMContentLoaded', async function() {
    const guestNameElement = document.querySelector('.guest-name');
    let guestNames = [];
    let index = 0;

    // SharePoint configuration
    const config = {
        auth: {
            clientId: 'd5d52da2-9c9f-4b98-b981-b886ea516f40',
            authority: 'https://login.microsoftonline.com/99496e88-1ae0-4a27-b6ac-47a10bb149dd',
            redirectUri: window.location.origin,
        }
    };

    const loginRequest = {
        scopes: ["https://graph.microsoft.com/.default"]
    };

    async function fetchExcelData() {
        try {
            const msalInstance = new msal.PublicClientApplication(config);
            let authResult;

            try {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    authResult = await msalInstance.acquireTokenSilent({
                        ...loginRequest,
                        account: accounts[0]
                    });
                } else {
                    authResult = await msalInstance.loginPopup(loginRequest);
                }
            } catch (loginError) {
                console.error('Login error:', loginError);
                throw loginError;
            }

            const filePath = '/Documents/EventWelcomer.xlsx:/RegisteredName';
            const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}`;
            
            const data = await fetch(endpoint, {
                headers: {
                    'Authorization': `Bearer ${authResult.accessToken}`
                }
            });
            
            const result = await data.json();
            guestNames = result.values.flat().filter(name => name);
        } catch (error) {
            console.error('Error fetching Excel data:', error);
            // Fallback to default names if fetch fails
            guestNames = ["Louisa", "John", "Emily", "Michael", "Sophia"];
        }
    }

    function changeGuestName() {
        guestNameElement.style.opacity = "0"; // Fade out
        setTimeout(() => {
            guestNameElement.textContent = guestNames[index];
            guestNameElement.style.opacity = "1"; // Fade in
            index = (index + 1) % guestNames.length;
        }, 500); // Delay before showing new name
    }

    // Initial fetch and start interval
    await fetchExcelData();
    setInterval(changeGuestName, 3000);
    setInterval(fetchExcelData, 300000); // Refresh data every 5 minutes
});
