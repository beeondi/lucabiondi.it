function checkNameParam() {
            const urlParams = new URLSearchParams(window.location.search);
            const nameParam = urlParams.get('name');
            
            if (nameParam) {
                const personalElement = document.getElementById('personal');
                if (personalElement) {
                    personalElement.style.display = 'none';
                }
                
                const nameElement = document.getElementById('name');
                if (nameElement) {
                    nameElement.textContent = nameParam;
                }
            }
        }

        function handleNameSubmit(event) {
            if (event.key === 'Enter') {
                const nameValue = event.target.value;
                if (nameValue.trim() !== '') {
                    const currentUrl = new URL(window.location.href);
                    currentUrl.searchParams.set('name', nameValue);
                    window.location.href = currentUrl.toString();
                }
            }
        }

        window.onload = checkNameParam;