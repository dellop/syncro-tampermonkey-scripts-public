// ==UserScript==
// @name         Syncro AI Ticket Creator
// @namespace    http://tampermonkey.net/
// @version      1.4
// @description  Create Syncro tickets from natural language descriptions using AI
// @author       mark
// @match        https://*.syncromsp.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      openrouter.ai
// @connect      syncromsp.com
// ==/UserScript==

(function() {
    'use strict';

    // Configuration - stored persistently
    let OPENROUTER_API_KEY = GM_getValue('openrouter_api_key', '');
    let SYNCRO_API_KEY = GM_getValue('syncro_api_key', '');
    let SYNCRO_SUBDOMAIN = GM_getValue('syncro_subdomain', '');
    let DEFAULT_AI_MODEL = GM_getValue('default_ai_model', '');
    let USE_SHIELD_DOMAIN = GM_getValue('use_shield_domain', true);

    // Store available models
    let availableModels = [];
    let syncroCustomers = [];
    let syncroUsers = [];
    let selectedCustomerId = null;
    let selectedUserId = null;

    // Store preloaded data
    let preloadedCustomers = [];
    let preloadedUsers = []; // Flat array of all users from all customers
    let isPreloading = false;
    let preloadCompleted = false;

    // Store parsed ticket data for review/submission
    let parsedTicketData = null;

    // Utility function for debouncing
    function debounce(func, delay) {
        let timeout;
        return function(...args) {
            const context = this;
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(context, args), delay);
        };
    }

    // Prompt for API keys if not set
    function promptForApiKeys() {
        if (!OPENROUTER_API_KEY || OPENROUTER_API_KEY === '') {
            const apiKey = prompt('Please enter your OpenRouter API key (get it from https://openrouter.ai):', '');
            if (apiKey && apiKey.trim()) {
                OPENROUTER_API_KEY = apiKey.trim();
                GM_setValue('openrouter_api_key', OPENROUTER_API_KEY);
                alert('OpenRouter API key saved successfully!');
            }
        }

        if (!SYNCRO_API_KEY || SYNCRO_API_KEY === '') {
            const apiKey = prompt('Please enter your Syncro API key:', '');
            if (apiKey && apiKey.trim()) {
                SYNCRO_API_KEY = apiKey.trim();
                GM_setValue('syncro_api_key', SYNCRO_API_KEY);
                alert('Syncro API key saved successfully!');
            }
        }
    }

    // Fetch available models from OpenRouter
    function fetchOpenRouterModels(callback) {
        GM_xmlhttpRequest({
            method: 'GET',
            url: 'https://openrouter.ai/api/v1/models',
            headers: {
                'Content-Type': 'application/json'
            },
            onload: function(response) {
                try {
                    const data = JSON.parse(response.responseText);
                    if (data.data) {
                        availableModels = data.data;
                        callback(availableModels);
                    } else {
                        console.error('Unexpected models response format:', data);
                        callback([]);
                    }
                } catch (e) {
                    console.error('Error parsing models response:', e);
                    callback([]);
                }
            },
            onerror: function(error) {
                console.error('Error fetching models:', error);
                callback([]);
            }
        });
    }

    // Fetch Syncro customers (handles pagination)
    function fetchSyncroCustomers(callback) {
        const allCustomers = [];

        function fetchPage(page) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://${SYNCRO_SUBDOMAIN}.syncromsp.com/api/v1/customers?api_key=${SYNCRO_API_KEY}&page=${page}`,
                headers: {
                    'Content-Type': 'application/json'
                },
                onload: function(response) {
                    try {
                        const data = JSON.parse(response.responseText);

                        if (data.customers && Array.isArray(data.customers)) {
                            // Add customers from this page
                            allCustomers.push(...data.customers);

                            // Check if there are more pages
                            const totalPages = data.meta ? data.meta.total_pages || 1 : 1;

                            if (page < totalPages) {
                                // Fetch next page
                                fetchPage(page + 1);
                            } else {
                                // All pages fetched, return all customers
                                syncroCustomers = allCustomers;
                                callback(null, syncroCustomers);
                            }
                        } else {
                            // No customers array, return what we have
                            syncroCustomers = allCustomers;
                            callback(null, syncroCustomers);
                        }
                    } catch (e) {
                        // On error, return what we have so far
                        callback('Error parsing customers response: ' + e.message, allCustomers);
                    }
                },
                onerror: function(error) {
                    // On network error, return what we have so far
                    callback('Error fetching customers: ' + JSON.stringify(error), allCustomers);
                }
            });
        }

        // Start fetching from page 1
        fetchPage(1);
    }

    // Fetch Syncro assets for a specific contact (user's computers) - handles pagination
    function fetchUserAssets(contactId, callback) {
        // First get the customer ID for this contact
        const contact = preloadedUsers.find(u => u.id == contactId);
        if (!contact) {
            callback('Contact not found', null);
            return;
        }

        const customerId = contact.customer_id;
        const allAssets = [];

        function fetchPage(page) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://${SYNCRO_SUBDOMAIN}.syncromsp.com/api/v1/customer_assets?customer_id=${customerId}&api_key=${SYNCRO_API_KEY}&page=${page}`,
                headers: {
                    'Content-Type': 'application/json'
                },
                onload: function(response) {
                    try {
                        const data = JSON.parse(response.responseText);

                        if (data.assets && Array.isArray(data.assets)) {
                            // Add assets from this page
                            allAssets.push(...data.assets);

                            // Debug: log the structure of the first asset to understand the data format
                            if (page === 1 && data.assets.length > 0) {
                                console.log('Sample asset structure:', data.assets[0]);
                            }

                            // Check if there are more pages
                            const totalPages = data.meta ? data.meta.total_pages || 1 : 1;

                            if (page < totalPages) {
                                // Fetch next page
                                fetchPage(page + 1);
                            } else {
                                // All pages fetched
                                // In this Syncro setup, assets are assigned to customers, not individual contacts
                                // So we show all customer assets when a user is selected
                                console.log('Assets are assigned to customers, not contacts. Showing all customer assets:', allAssets.length);
                                callback(null, allAssets);
                            }
                        } else {
                            // No assets array, return what we have
                            callback(null, allAssets);
                        }
                    } catch (e) {
                        // On error, return what we have so far
                        callback('Error parsing assets response: ' + e.message, allAssets);
                    }
                },
                onerror: function(error) {
                    // On network error, return what we have so far
                    callback('Error fetching user assets: ' + JSON.stringify(error), allAssets);
                }
            });
        }

        // Start fetching from page 1
        fetchPage(1);
    }

    // Fetch Syncro users for a specific customer (handles pagination)
    function fetchSyncroUsers(customerId, callback) {
        const allContacts = [];

        function fetchPage(page) {
            GM_xmlhttpRequest({
                method: 'GET',
                url: `https://${SYNCRO_SUBDOMAIN}.syncromsp.com/api/v1/contacts?customer_id=${customerId}&page=${page}&api_key=${SYNCRO_API_KEY}`,
                headers: {
                    'Content-Type': 'application/json'
                },
                onload: function(response) {
                    try {
                        const data = JSON.parse(response.responseText);

                        if (data.contacts && Array.isArray(data.contacts)) {
                            // Add contacts from this page
                            allContacts.push(...data.contacts);

                            // Check if there are more pages
                            const totalPages = data.meta ? data.meta.total_pages || 1 : 1;

                            if (page < totalPages) {
                                // Fetch next page
                                fetchPage(page + 1);
                            } else {
                                // All pages fetched, return all contacts
                                callback(null, allContacts);
                            }
                        } else {
                            // No contacts array, return what we have
                            callback(null, allContacts);
                        }
                    } catch (e) {
                        // On error, return what we have so far
                        callback(null, allContacts);
                    }
                },
                onerror: function(error) {
                    // On network error, return what we have so far
                    callback(null, allContacts);
                }
            });
        }

        // Start fetching from page 1
        fetchPage(1);
    }

    // Function to fetch all users for all preloaded customers
    function fetchAllUsersForCustomers(customers, callback) {
        const userFetchPromises = customers.map(customer => {
            return new Promise(resolve => {
                fetchSyncroUsers(customer.id, (error, users) => {
                    if (!error && users) {
                        // Add customer_id to each user for easier lookup later
                        users.forEach(user => user.customer_id = customer.id);
                        resolve(users);
                    } else {
                        console.warn(`Could not fetch users for customer ${customer.id}: ${error}`);
                        resolve([]); // Resolve with empty array on error
                    }
                });
            });
        });

        Promise.all(userFetchPromises)
            .then(results => {
                preloadedUsers = results.flat(); // Flatten all user arrays into one
                console.log(`Preloaded ${preloadedUsers.length} users.`);
                callback(null, preloadedUsers);
            })
            .catch(error => {
                console.error('Error fetching all users for customers:', error);
                callback(error, []);
            });
    }

    // Function to start preloading organizations and users
    function startPreloadingData() {
        if (isPreloading || preloadCompleted) {
            console.log('Preloading already in progress or completed.');
            return;
        }

        isPreloading = true;
        console.log('Starting background preload of organizations and users...');

        // Display a loading message in the sidebar if it exists
        const resultDiv = document.getElementById('ticket-creator-result');
        if (resultDiv) {
            resultDiv.innerHTML = '<div class="ticket-creator-result info">Loading organizations and users in background...</div>';
        }

        fetchSyncroCustomers((error, customers) => {
            if (error) {
                console.error('Error preloading customers:', error);
                isPreloading = false;
                if (resultDiv) {
                    resultDiv.innerHTML = '<div class="ticket-creator-result error">Error preloading organizations.</div>';
                }
                return;
            }

            preloadedCustomers = customers;
            console.log(`Preloaded ${preloadedCustomers.length} organizations.`);

            fetchAllUsersForCustomers(preloadedCustomers, (userError, users) => {
                isPreloading = false;
                if (userError) {
                    console.error('Error preloading users:', userError);
                    if (resultDiv) {
                        resultDiv.innerHTML = '<div class="ticket-creator-result error">Error preloading users.</div>';
                    }
                    return;
                }

                preloadCompleted = true;
                console.log('Preloading of organizations and users completed.');
                if (resultDiv && resultDiv.innerHTML.includes('Loading organizations and users')) {
                    // Clear the loading message if it's still there
                    resultDiv.innerHTML = '';
                }
            });
        });
    }

    // Valid problem types according to Syncro API requirements
    const VALID_PROBLEM_TYPES = [
        'Hardware',
        'Software',
        'Project / Planned Work',
        'Network / Connectivity',
        'New Device / Deployment',
        'Maintenance / Preventitive',
        'User Account / Access',
        'Security / Malware',
        'Internal / MSP Operations',
        'Other'
    ];

    // Validate and normalize problem type
    function validateProblemType(problemType) {
        if (!problemType) return 'Other';

        // Check for exact match (case-insensitive)
        const normalized = VALID_PROBLEM_TYPES.find(
            validType => validType.toLowerCase() === problemType.toLowerCase()
        );

        if (normalized) return normalized;

        // If no exact match, try partial matching for common variations
        const lowerProblemType = problemType.toLowerCase();

        if (lowerProblemType.includes('hardware') || lowerProblemType.includes('device issue')) {
            return 'Hardware';
        }
        if (lowerProblemType.includes('software') || lowerProblemType.includes('application')) {
            return 'Software';
        }
        if (lowerProblemType.includes('network') || lowerProblemType.includes('connectivity') || lowerProblemType.includes('internet')) {
            return 'Network / Connectivity';
        }
        if (lowerProblemType.includes('project') || lowerProblemType.includes('planned')) {
            return 'Project / Planned Work';
        }
        if (lowerProblemType.includes('deployment') || lowerProblemType.includes('new device') || lowerProblemType.includes('setup')) {
            return 'New Device / Deployment';
        }
        if (lowerProblemType.includes('maintenance') || lowerProblemType.includes('preventive') || lowerProblemType.includes('preventitive')) {
            return 'Maintenance / Preventitive';
        }
        if (lowerProblemType.includes('account') || lowerProblemType.includes('access') || lowerProblemType.includes('password') || lowerProblemType.includes('login')) {
            return 'User Account / Access';
        }
        if (lowerProblemType.includes('security') || lowerProblemType.includes('malware') || lowerProblemType.includes('virus')) {
            return 'Security / Malware';
        }
        if (lowerProblemType.includes('MSP') || lowerProblemType.includes('internal') || lowerProblemType.includes('operations')) {
            return 'Internal / MSP Operations';
        }

        // Default to 'Other' if no match found
        return 'Other';
    }

    // Fallback function to extract names from description if AI fails
    function extractNamesFromDescription(description) {
        const names = [];
        // Match common name patterns: word starting with capital letter, optionally followed by lowercase letters
        // Look for patterns like "John called", "Sarah from", "Mike reported", etc.
        const namePatterns = [
            /\b([A-Z][a-z]+)\s+(called|said|reported|from|at|needs?|has|is having|experienced?)\b/g,
            /\b([A-Z][a-z]+)\s+([A-Z][a-z]+)?\s+(called|said|reported|needs?|has|is having|experienced?)\b/g,
            /\b([A-Z][a-z]+)\s+([A-Z][a-z]+)\b/g  // Simple first + last name pattern
        ];

        namePatterns.forEach(pattern => {
            let match;
            while ((match = pattern.exec(description)) !== null) {
                const name = match[1] + (match[2] ? ' ' + match[2] : '');
                if (!names.includes(name) && name.length > 1) {  // Avoid single letters
                    names.push(name);
                }
            }
        });

        return names;
    }

    // Call OpenRouter API to parse ticket description
    function parseTicketDescription(description, model, callback) {
        if (!OPENROUTER_API_KEY || OPENROUTER_API_KEY === '') {
            callback('Error: OpenRouter API key is required.', null);
            return;
        }

        const prompt = `You are a ticket parsing assistant. Parse the following ticket description and extract:
1. Customer/Organization name (company name) - if not mentioned, use empty string ""
2. User name (person's name) - extract any person mentioned who is reporting or experiencing the issue
3. Computer reference - TRUE if the description mentions "their computer", "my computer", "the computer", "his computer", "her computer", "laptop", "desktop", "workstation", or similar device references. FALSE otherwise.
4. A concise subject line for the ticket (max 80 characters)
5. The initial issue description (cleaned up and professional, without including any user or company names)
6. The problem type/category - MUST be EXACTLY one of these values:
   - "Hardware"
   - "Software"
   - "Project / Planned Work"
   - "Network / Connectivity"
   - "New Device / Deployment"
   - "Maintenance / Preventitive"
   - "User Account / Access"
   - "Security / Malware"
   - "Internal / MSP Operations"
   - "Other"

Examples:
- "John called about his computer not working" → user: "John", organization: "", computer_reference: true
- "Sarah from ABC Corp said her email isn't working" → user: "Sarah", organization: "ABC Corp", computer_reference: false
- "Mike reported that drives aren't mapping on his laptop" → user: "Mike", organization: "", computer_reference: true
- "The server at XYZ Company is down" → user: "", organization: "XYZ Company", computer_reference: false

Ticket description:
"${description}"

Respond ONLY with valid JSON in this exact format (no markdown, no code blocks, just raw JSON):
{
  "organization": "extracted organization name or empty string",
  "user": "extracted user name or empty string",
  "computer_reference": true or false,
  "subject": "concise subject line",
  "issue": "cleaned up issue description",
  "problem_type": "problem category (must be one of the exact values listed above)"
}`;

        const requestBody = {
            model: model,
            messages: [
                {
                    role: 'user',
                    content: prompt
                }
            ],
            temperature: 0.3,
            max_tokens: 500
        };

        GM_xmlhttpRequest({
            method: 'POST',
            url: 'https://openrouter.ai/api/v1/chat/completions',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
                'HTTP-Referer': window.location.href,
                'X-Title': 'Syncro Ticket Creator'
            },
            data: JSON.stringify(requestBody),
            onload: function(response) {
                try {
                    const data = JSON.parse(response.responseText);
                    if (data.choices && data.choices[0] && data.choices[0].message) {
                        const content = data.choices[0].message.content.trim();
                        // Remove markdown code blocks if present
                        const cleanContent = content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
                        const parsed = JSON.parse(cleanContent);
                        // Validate and normalize problem_type to ensure it matches one of the valid values
                        if (parsed.problem_type) {
                            parsed.problem_type = validateProblemType(parsed.problem_type);
                        } else {
                            parsed.problem_type = 'Other';
                        }
                        callback(null, parsed);
                    } else {
                        callback('Unexpected response format from OpenRouter', null);
                    }
                } catch (e) {
                    callback('Error parsing OpenRouter response: ' + e.message, null);
                }
            },
            onerror: function(error) {
                callback('Error calling OpenRouter API: ' + JSON.stringify(error), null);
            }
        });
    }

    // Find customer ID by name (now uses preloadedCustomers)
    function findCustomerByName(customerName) {
        const lowerName = customerName.toLowerCase();
        return preloadedCustomers.find(c =>
            c.business_name && c.business_name.toLowerCase().includes(lowerName) ||
            c.firstname && c.lastname && `${c.firstname} ${c.lastname}`.toLowerCase().includes(lowerName)
        );
    }

    // Find user by name within customer
    function findUserByName(users, userName) {
        if (!userName || !userName.trim()) return null;

        const searchName = userName.toLowerCase().trim();

        // Debug: log user data structure
        if (users.length > 0) {
            console.log('Sample user data structure:', users[0]);
        }

        return users.find(u => {
            // Get all available name fields and convert to lowercase
            const fullName = u.name ? u.name.toLowerCase() : '';
            const firstName = u.firstname ? u.firstname.toLowerCase() : '';
            const lastName = u.lastname ? u.lastname.toLowerCase() : '';
            const combinedName = (firstName && lastName) ? `${firstName} ${lastName}` : '';

            // Check if search name appears in any name field
            if (fullName && fullName.includes(searchName)) return true;
            if (combinedName && combinedName.includes(searchName)) return true;
            if (firstName && firstName.includes(searchName)) return true;
            if (lastName && lastName.includes(searchName)) return true;

            return false;
        });
    }

    // Search for contact via Syncro search endpoint for faster lookup
    function searchForContact(userName, callback) {
        GM_xmlhttpRequest({
            method: 'GET',
            url: `https://${SYNCRO_SUBDOMAIN}.syncromsp.com/api/v1/search?query=${encodeURIComponent(userName)}&api_key=${SYNCRO_API_KEY}`,
            headers: { 'Content-Type': 'application/json' },
            onload: function(response) {
                try {
                    const data = JSON.parse(response.responseText);
                    const results = data.results || [];
                    const contacts = results
                        .filter(r => r.table && r.table._type === 'contact')
                        .map(r => r.table._source.table);
                    callback(null, contacts);
                } catch (e) {
                    callback(e, []);
                }
            },
            onerror: function(err) {
                callback(err, []);
            }
        });
    }

    // Find user across all customers and return customer info (now uses preloaded data)
    function findUserAcrossAllCustomers(userName, callback) {
        if (!preloadCompleted) {
            // If preloading isn't complete, fall back to original behavior or indicate error
            // For now, we'll just return an empty array and let the UI handle it
            console.warn('Preloading not completed, cannot search preloaded users.');
            callback(null, []);
            return;
        }

        const searchName = userName.toLowerCase().trim();
        const foundUsers = [];

        preloadedUsers.forEach(user => {
            const fullName = user.name ? user.name.toLowerCase() : '';
            const firstName = user.firstname ? user.firstname.toLowerCase() : '';
            const lastName = user.lastname ? user.lastname.toLowerCase() : '';
            const combinedName = (firstName && lastName) ? `${firstName} ${lastName}` : '';

            if ((fullName && fullName.includes(searchName)) ||
                (combinedName && combinedName.includes(searchName)) ||
                (firstName && firstName.includes(searchName)) ||
                (lastName && lastName.includes(searchName))) {

                const customer = preloadedCustomers.find(c => c.id === user.customer_id);
                if (customer) {
                    foundUsers.push({ user: user, customer: customer });
                }
            }
        });

        callback(null, foundUsers);
    }

    // Populate user dropdown with users from selected organization
    function populateUserDropdown(users, selectedUserName) {
        const userSelect = document.getElementById('parsed-user');
        if (!userSelect) return;

        userSelect.innerHTML = '<option value="">Select User...</option>';

        users.forEach(user => {
            const option = document.createElement('option');
            option.value = user.id;
            option.textContent = user.name || `${user.firstname} ${user.lastname}`;
            // Try to match the parsed user name
            if (selectedUserName &&
                (user.name && user.name.toLowerCase().includes(selectedUserName.toLowerCase()) ||
                 `${user.firstname} ${user.lastname}`.toLowerCase().includes(selectedUserName.toLowerCase()))) {
                option.selected = true;
                selectedUserId = user.id;
            }
            userSelect.appendChild(option);
        });

        // If a user was auto-selected, load their computers
        if (selectedUserId) {
            handleUserChange(selectedUserId);
        }
    }

    // Populate computer dropdown with user's assets
    function populateComputerDropdown(assets, autoSelectSingle = false) {
        const computerSection = document.getElementById('computer-selection-section');
        const computerSelect = document.getElementById('parsed-computer');

        if (!computerSection || !computerSelect) return;

        computerSelect.innerHTML = '<option value="">No computer selected</option>';

        if (assets && assets.length > 0) {
            // Sort assets alphabetically by name
            const sortedAssets = assets.slice().sort((a, b) => {
                const nameA = (a.name || `Asset #${a.id}`).toLowerCase();
                const nameB = (b.name || `Asset #${b.id}`).toLowerCase();
                return nameA.localeCompare(nameB);
            });

            sortedAssets.forEach(asset => {
                const option = document.createElement('option');
                option.value = asset.id;
                option.textContent = asset.name || `Asset #${asset.id}`;
                computerSelect.appendChild(option);
            });

            // Show the computer section
            computerSection.style.display = 'block';

            // Auto-select if user mentioned computer and has exactly one
            if (autoSelectSingle && assets.length === 1 && parsedTicketData && parsedTicketData.computer_reference) {
                computerSelect.value = assets[0].id;
            }
        } else {
            // Hide the computer section if no assets
            computerSection.style.display = 'none';
        }
    }

    // Handle organization selection change
    function handleOrganizationChange(event) {
        const customerId = event.target.value;
        selectedCustomerId = customerId;
        selectedUserId = null;

        if (customerId) {
            // Use preloaded users for the selected customer
            const usersForCustomer = preloadedUsers.filter(user => user.customer_id == customerId);
            populateUserDropdown(usersForCustomer, parsedTicketData ? parsedTicketData.user : '');
        } else {
            // Clear user dropdown
            const userSelect = document.getElementById('parsed-user');
            if (userSelect) {
                userSelect.innerHTML = '<option value="">Select an organization first</option>';
            }
        }

        // Clear computer selection when organization changes
        populateComputerDropdown([]);
    }

    // Handle user selection change
    function handleUserChange(userId) {
        selectedUserId = userId;

        if (userId) {
            // Fetch and populate user's computers
            fetchUserAssets(userId, (error, assets) => {
                if (error) {
                    console.error('Error fetching user assets:', error);
                    populateComputerDropdown([]);
                } else {
                    console.log('Fetched assets for user', userId, ':', assets);
                    // Auto-select computer if user mentioned computer and has exactly one
                    const autoSelect = parsedTicketData && parsedTicketData.computer_reference && assets.length === 1;
                    populateComputerDropdown(assets, autoSelect);
                }
            });
        } else {
            // Clear computer selection
            populateComputerDropdown([]);
        }
    }

    // Populate dropdowns when we have a specific customer and user
    function populateDropdownsWithCustomer(customer, user, parsed) {
        const orgSelect = document.getElementById('parsed-organization');
        if (!orgSelect) return;

        orgSelect.innerHTML = '<option value="">Select Organization...</option>';

        // Add the found customer
        const option = document.createElement('option');
        option.value = customer.id;
        option.textContent = customer.business_name || `${customer.firstname} ${customer.lastname}`;
        option.selected = true;
        selectedCustomerId = customer.id;
        orgSelect.appendChild(option);

        // Add event listener for organization change
        orgSelect.addEventListener('change', handleOrganizationChange);

        // Add event listener for user change
        const userSelect = document.getElementById('parsed-user');
        if (userSelect) {
            userSelect.addEventListener('change', (event) => {
                handleUserChange(event.target.value);
            });
        }

        // Populate user dropdown
        populateUserDropdown([user], parsed.user);

        // Populate other fields
        const subjectInput = document.getElementById('parsed-subject');
        if (subjectInput) subjectInput.value = parsed.subject || '';

        const issueTextarea = document.getElementById('parsed-issue');
        if (issueTextarea) issueTextarea.value = parsed.issue || '';

        const problemTypeSelect = document.getElementById('parsed-problem-type');
        if (problemTypeSelect) problemTypeSelect.value = parsed.problem_type || 'Other';

        // Show review section and hide initial form
        const reviewSection = document.getElementById('ticket-review-section');
        if (reviewSection) reviewSection.style.display = 'block';

        const inputSections = document.querySelectorAll('.ticket-creator-sidebar-content > .ticket-creator-input-section');
        inputSections.forEach(section => {
            if (section) section.style.display = 'none';
        });

        const submitBtn = document.getElementById('ticket-creator-submit-btn');
        if (submitBtn) submitBtn.style.display = 'none';

        const settingsBtn = document.getElementById('ticket-creator-settings-btn');
        if (settingsBtn) settingsBtn.style.display = 'none';

        const resultDiv = document.getElementById('ticket-creator-result');
        if (resultDiv) {
            resultDiv.innerHTML = '<div class="ticket-creator-result success">✓ Information parsed successfully! Please review and edit the details below, then click "Submit Ticket".</div>';
        }
    }

    // Populate dropdowns with all customers (normal flow)
    function populateDropdownsWithCustomers(customers, parsed) {
        const orgSelect = document.getElementById('parsed-organization');
        if (!orgSelect) return;

        orgSelect.innerHTML = '<option value="">Select Organization...</option>';

        customers.forEach(customer => {
            const option = document.createElement('option');
            option.value = customer.id;
            option.textContent = customer.business_name || `${customer.firstname} ${customer.lastname}`;
            // Try to match the parsed organization name
            if (parsed.organization &&
                (customer.business_name && customer.business_name.toLowerCase().includes(parsed.organization.toLowerCase()) ||
                 `${customer.firstname} ${customer.lastname}`.toLowerCase().includes(parsed.organization.toLowerCase()))) {
                option.selected = true;
                selectedCustomerId = customer.id;
                // Use preloaded users for this customer
                const usersForCustomer = preloadedUsers.filter(user => user.customer_id == customer.id);
                populateUserDropdown(usersForCustomer, parsed.user);
            }
            orgSelect.appendChild(option);
        });

        // Add event listener for organization change
        orgSelect.addEventListener('change', handleOrganizationChange);

        // Populate other fields
        const subjectInput = document.getElementById('parsed-subject');
        if (subjectInput) subjectInput.value = parsed.subject || '';

        const issueTextarea = document.getElementById('parsed-issue');
        if (issueTextarea) issueTextarea.value = parsed.issue || '';

        const problemTypeSelect = document.getElementById('parsed-problem-type');
        if (problemTypeSelect) problemTypeSelect.value = parsed.problem_type || 'Other';

        // Show review section and hide initial form
        const reviewSection = document.getElementById('ticket-review-section');
        if (reviewSection) reviewSection.style.display = 'block';

        const inputSections = document.querySelectorAll('.ticket-creator-sidebar-content > .ticket-creator-input-section');
        inputSections.forEach(section => {
            if (section) section.style.display = 'none';
        });

        const submitBtn = document.getElementById('ticket-creator-submit-btn');
        if (submitBtn) {
            submitBtn.style.display = 'none';
            submitBtn.disabled = false;
            submitBtn.textContent = 'Send';
        }

        const settingsBtn = document.getElementById('ticket-creator-settings-btn');
        if (settingsBtn) settingsBtn.style.display = 'none';

        const resultDiv = document.getElementById('ticket-creator-result');
        if (resultDiv) {
            resultDiv.innerHTML = '<div class="ticket-creator-result success">✓ Information parsed successfully! Please review and edit the details below, then click "Submit Ticket".</div>';
        }
    }

    // Display options for user selection when multiple users are found
    function displayUserSelectionOptions(foundUsers, parsedUserName) {
        const resultDiv = document.getElementById('ticket-creator-result');
        const submitBtn = document.getElementById('ticket-creator-submit-btn');
        const inputSections = document.querySelectorAll('.ticket-creator-sidebar-content > .ticket-creator-input-section');

        // Hide initial input sections and submit button
        inputSections.forEach(section => {
            if (section) section.style.display = 'none';
        });
        if (submitBtn) submitBtn.style.display = 'none';

        let optionsHtml = `<div class="ticket-creator-result info">Multiple users named "${parsedUserName}" were found. Please select the correct user:</div>`;
        optionsHtml += '<div style="margin-top: 15px;">';

        foundUsers.forEach((item, index) => {
            const userName = item.user.name || `${item.user.firstname} ${item.user.lastname}`;
            const customerName = item.customer.business_name || `${item.customer.firstname} ${item.customer.lastname}`;
            optionsHtml += `
                <button class="ticket-creator-select-user-btn" data-user-index="${index}"
                        style="width: 100%; background: #007bff; color: white; border: none; padding: 10px 16px; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: 500; margin-bottom: 8px;">
                    ${userName} (${customerName})
                </button>
            `;
        });
        optionsHtml += '</div>';

        resultDiv.innerHTML = optionsHtml;

        // Add event listeners to the new buttons
        document.querySelectorAll('.ticket-creator-select-user-btn').forEach(button => {
            button.addEventListener('click', (e) => {
                const index = parseInt(e.target.dataset.userIndex);
                const selectedItem = foundUsers[index];

                // Update parsedTicketData with the selected user and customer
                parsedTicketData.organization = selectedItem.customer.business_name || `${selectedItem.customer.firstname} ${selectedItem.customer.lastname}`;
                parsedTicketData.user = selectedItem.user.name || `${selectedItem.user.firstname} ${selectedItem.user.lastname}`;

                // Proceed with populating dropdowns and showing review section
                populateDropdownsWithCustomer(selectedItem.customer, selectedItem.user, parsedTicketData);
            });
        });
    }

    // Create Syncro ticket
    function createSyncroTicket(ticketData, callback) {
        const requestBody = {
            customer_id: ticketData.customer_id,
            subject: ticketData.subject,
            problem_type: ticketData.problem_type,
            status: 'New',
            comments_attributes: [{
                subject: 'Initial Issue',
                body: ticketData.body,
                hidden: false,
                do_not_email: !ticketData.send_email  // Use checkbox value for email notifications
            }]
        };

        if (ticketData.user_id) {
            requestBody.contact_id = ticketData.user_id;
        }

        if (ticketData.asset_ids && ticketData.asset_ids.length > 0) {
            requestBody.asset_ids = ticketData.asset_ids;
        }

        GM_xmlhttpRequest({
            method: 'POST',
            url: `https://${SYNCRO_SUBDOMAIN}.syncromsp.com/api/v1/tickets?api_key=${SYNCRO_API_KEY}`,
            headers: {
                'Content-Type': 'application/json'
            },
            data: JSON.stringify(requestBody),
            onload: function(response) {
                try {
                    const data = JSON.parse(response.responseText);
                    if (data.ticket && data.ticket.id) {
                        callback(null, data.ticket);
                    } else if (response.status >= 200 && response.status < 300) {
                        // Success but different response format
                        callback(null, data);
                    } else {
                        callback('Failed to create ticket: ' + response.responseText, null);
                    }
                } catch (e) {
                    callback('Error parsing create ticket response: ' + e.message, null);
                }
            },
            onerror: function(error) {
                callback('Error creating ticket: ' + JSON.stringify(error), null);
            }
        });
    }

    // Create shared styles
    function createSharedStyles() {
        const style = document.createElement('style');
        style.textContent = `
            /* Ticket Creator sidebar styles */
            .ticket-creator-sidebar {
                position: fixed;
                left: 0;
                top: 0;
                width: 450px;
                height: 100vh;
                background: white;
                box-shadow: 2px 0 10px rgba(0,0,0,0.1);
                z-index: 10000;
                display: flex;
                flex-direction: column;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                transition: transform 0.3s ease, width 0.3s ease;
                resize: horizontal;
                overflow: auto;
                min-width: 350px;
                max-width: 800px;
            }

            .ticket-creator-sidebar-resizer {
                position: absolute;
                top: 0;
                right: 0;
                width: 8px; /* Make it a bit wider for easier grabbing */
                height: 100%;
                cursor: ew-resize;
                z-index: 10001;
            }

            .ticket-creator-sidebar.resizing {
                transition: none;
            }

            .ticket-creator-sidebar.minimized {
                transform: translateX(-450px);
            }

            /* Push page content when sidebar is open */
            body.ticket-creator-sidebar-open {
                margin-left: 450px;
                transition: margin-left 0.3s ease;
            }

            .ticket-creator-sidebar-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 20px;
                border-bottom: 1px solid #e0e0e0;
                background: linear-gradient(135deg, #34a853 0%, #2d8e47 100%);
            }

            .ticket-creator-header-buttons {
                display: flex;
                gap: 8px;
            }

            .ticket-creator-sidebar-header h3 {
                margin: 0;
                font-size: 18px;
                color: white;
            }

            .ticket-creator-close-btn {
                background: none;
                border: none;
                font-size: 24px;
                cursor: pointer;
                padding: 0;
                width: 30px;
                height: 30px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 4px;
                transition: background 0.2s;
                color: white;
            }

            .ticket-creator-close-btn:hover {
                background: rgba(255,255,255,0.2);
            }

            .ticket-creator-settings-icon-btn {
                background: none;
                border: none;
                font-size: 18px;
                cursor: pointer;
                padding: 0;
                width: 30px;
                height: 30px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 4px;
                transition: background 0.2s;
                color: white;
            }

            .ticket-creator-settings-icon-btn:hover {
                background: rgba(255,255,255,0.2);
            }

            .ticket-creator-sidebar-content {
                flex: 1;
                overflow-y: auto;
                padding: 20px;
            }

            .ticket-creator-input-section {
                margin-bottom: 20px;
            }

            .ticket-creator-input-section label {
                display: block;
                margin-bottom: 8px;
                color: #333;
                font-weight: 500;
                font-size: 13px;
            }

            .ticket-creator-input-section input,
            .ticket-creator-input-section select,
            .ticket-creator-input-section textarea {
                width: 100%;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 4px;
                font-family: inherit;
                font-size: 13px;
                box-sizing: border-box;
                color: #333;
                background: white;
            }

            .ticket-creator-input-section textarea {
                min-height: 120px;
                resize: vertical;
            }

            .ticket-creator-input-section input:focus,
            .ticket-creator-input-section select:focus,
            .ticket-creator-input-section textarea:focus {
                outline: none;
                border-color: #34a853;
            }

            .ticket-creator-submit-btn {
                width: 100%;
                background: #34a853;
                color: white;
                border: none;
                padding: 12px 16px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                font-weight: 500;
                margin-top: 10px;
            }

            .ticket-creator-submit-btn:hover {
                background: #2d8e47;
            }

            .ticket-creator-submit-btn:disabled {
                background: #ccc;
                cursor: not-allowed;
            }

            .ticket-creator-settings-btn {
                width: 100%;
                background: #667eea;
                color: white;
                border: none;
                padding: 10px 16px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 13px;
                font-weight: 500;
                margin-top: 10px;
            }

            .ticket-creator-settings-btn:hover {
                background: #5568d3;
            }

            .ticket-creator-restore-tab {
                position: fixed;
                left: 0;
                top: calc(50% + 180px);
                background: linear-gradient(135deg, #34a853 0%, #2d8e47 100%);
                color: white;
                border: none;
                padding: 10px 5px;
                cursor: pointer;
                z-index: 9999;
                writing-mode: vertical-rl;
                text-orientation: mixed;
                border-radius: 0 5px 5px 0;
                font-size: 12px;
                font-weight: 500;
                box-shadow: 2px 0 5px rgba(0,0,0,0.2);
                display: block;
            }

            .ticket-creator-restore-tab:hover {
                background: linear-gradient(135deg, #2d8e47 0%, #34a853 100%);
            }

            .ticket-creator-result {
                margin-top: 20px;
                padding: 15px;
                border-radius: 4px;
                font-size: 13px;
                line-height: 1.5;
            }

            .ticket-creator-result.success {
                background: #d4edda;
                border: 1px solid #c3e6cb;
                color: #155724;
            }

            .ticket-creator-result.error {
                background: #f8d7da;
                border: 1px solid #f5c6cb;
                color: #721c24;
            }

            .ticket-creator-result.info {
                background: #d1ecf1;
                border: 1px solid #bee5eb;
                color: #0c5460;
            }

            .ticket-creator-loading {
                text-align: center;
                padding: 20px;
                color: #666;
            }

            .ticket-creator-config-modal {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                z-index: 10001;
                max-width: 500px;
                width: 90%;
            }

            .ticket-creator-modal-overlay {
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: rgba(0,0,0,0.5);
                z-index: 10000;
            }

            .ticket-creator-config-modal h3 {
                margin-top: 0;
                color: #333;
            }

            .ticket-creator-config-modal input {
                width: 100%;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 4px;
                margin-bottom: 15px;
                font-size: 13px;
                box-sizing: border-box;
            }

            .ticket-creator-config-modal button {
                padding: 10px 20px;
                margin-right: 10px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 13px;
                font-weight: 500;
            }

            .ticket-creator-config-modal .save-btn {
                background: #34a853;
                color: white;
            }

            .ticket-creator-config-modal .cancel-btn {
                background: #ccc;
                color: #333;
            }
        `;
        document.head.appendChild(style);
    }

    // Show configuration modal
    function showConfigModal() {
        // Remove existing modal if any
        const existingOverlay = document.querySelector('.ticket-creator-modal-overlay');
        if (existingOverlay) existingOverlay.remove();

        const existingModal = document.querySelector('.ticket-creator-config-modal');
        if (existingModal) existingModal.remove();

        // Create overlay
        const overlay = document.createElement('div');
        overlay.className = 'ticket-creator-modal-overlay';
        document.body.appendChild(overlay);

        // Create modal
        const modal = document.createElement('div');
        modal.className = 'ticket-creator-config-modal';
        modal.innerHTML = `
            <h3>API Configuration</h3>
            <label style="display: block; margin-bottom: 5px; font-weight: 500; font-size: 13px;">OpenRouter API Key:</label>
            <input type="password" id="config-openrouter-key" value="${OPENROUTER_API_KEY}" placeholder="sk-or-...">

            <label style="display: block; margin-bottom: 5px; font-weight: 500; font-size: 13px;">Syncro API Key:</label>
            <input type="password" id="config-syncro-key" value="${SYNCRO_API_KEY}" placeholder="Your Syncro API key">

            <label style="display: block; margin-bottom: 5px; font-weight: 500; font-size: 13px;">Syncro Subdomain:</label>
            <input type="text" id="config-syncro-subdomain" value="${SYNCRO_SUBDOMAIN}" placeholder="e.g., firebytes">

            <label style="display: block; margin-bottom: 5px; font-weight: 500; font-size: 13px;">Default AI Model:</label>
            <select id="config-default-model">
                <option value="">Select default model...</option>
            </select>

            <label style="display: block; margin-bottom: 5px; font-weight: 500; font-size: 13px;">Domain:</label>
            <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 15px;">
                <input type="checkbox" id="config-use-shield-domain" ${USE_SHIELD_DOMAIN ? 'checked' : ''} style="margin: 0;">
                Use shield.syncromsp.com domain
            </label>

            <div style="margin-top: 20px;">
                <button class="save-btn" id="config-save-btn">Save</button>
                <button class="cancel-btn" id="config-cancel-btn">Cancel</button>
            </div>
        `;
        document.body.appendChild(modal);

        // Populate default model dropdown
        fetchOpenRouterModels((models) => {
            const modelSelect = document.getElementById('config-default-model');
            if (models.length > 0) {
                // Sort models alphabetically by name
                const sortedModels = models.sort((a, b) => a.name.localeCompare(b.name));

                sortedModels.forEach(model => {
                    const option = document.createElement('option');
                    option.value = model.id;
                    option.textContent = model.name;
                    if (model.id === DEFAULT_AI_MODEL) {
                        option.selected = true;
                    }
                    modelSelect.appendChild(option);
                });
            }
        });

        // Handle save
        document.getElementById('config-save-btn').addEventListener('click', () => {
            OPENROUTER_API_KEY = document.getElementById('config-openrouter-key').value.trim();
            SYNCRO_API_KEY = document.getElementById('config-syncro-key').value.trim();
            SYNCRO_SUBDOMAIN = document.getElementById('config-syncro-subdomain').value.trim();
            DEFAULT_AI_MODEL = document.getElementById('config-default-model').value.trim();
            USE_SHIELD_DOMAIN = document.getElementById('config-use-shield-domain').checked;

            GM_setValue('openrouter_api_key', OPENROUTER_API_KEY);
            GM_setValue('syncro_api_key', SYNCRO_API_KEY);
            GM_setValue('syncro_subdomain', SYNCRO_SUBDOMAIN);
            GM_setValue('default_ai_model', DEFAULT_AI_MODEL);
            GM_setValue('use_shield_domain', USE_SHIELD_DOMAIN);

            overlay.remove();
            modal.remove();

            alert('Configuration saved successfully!');
        });

        // Handle cancel
        document.getElementById('config-cancel-btn').addEventListener('click', () => {
            overlay.remove();
            modal.remove();
        });

        // Close on overlay click
        overlay.addEventListener('click', () => {
            overlay.remove();
            modal.remove();
        });
    }

    // Create Ticket Creator Sidebar
    function createTicketCreatorSidebar() {
        const existingSidebar = document.getElementById('ticket-creator-sidebar');
        if (existingSidebar) {
            existingSidebar.remove();
        }

        const sidebar = document.createElement('div');
        sidebar.id = 'ticket-creator-sidebar';
        sidebar.className = 'ticket-creator-sidebar minimized';
        sidebar.innerHTML = `
            <div class="ticket-creator-sidebar-header">
                <h3>🎫 Create Ticket</h3>
                <div class="ticket-creator-header-buttons">
                    <button class="ticket-creator-settings-icon-btn" id="ticket-creator-settings-icon-btn" title="API Settings">⚙️</button>
                    <button class="ticket-creator-close-btn" id="ticket-creator-close-btn" title="Close">&times;</button>
                </div>
            </div>
            <div class="ticket-creator-sidebar-content">
                <div class="ticket-creator-input-section">
                    <label for="ticket-description-input">Ticket Description:</label>
                    <textarea id="ticket-description-input" placeholder="e.g., 'Trever from Bollig called and his drive letters aren't mapping on his microsoft surface.'"></textarea>
                    <small style="color: #666; font-size: 11px; display: block; margin-top: 5px;">
                        Include the person's name, company name, and the issue description. The AI will extract the details automatically.
                    </small>
                </div>

                <div class="ticket-creator-input-section">
                    <label for="ticket-model-select">AI Model:</label>
                    <select id="ticket-model-select">
                        <option value="">Loading models...</option>
                    </select>
                </div>

                <button class="ticket-creator-submit-btn" id="ticket-creator-submit-btn">Send</button>

                <!-- Review/Edit Section (initially hidden) -->
                <div id="ticket-review-section" style="display: none;">
                    <h4 style="margin: 20px 0 15px 0; color: #333; font-size: 14px;">Review & Edit Parsed Information:</h4>

                    <div class="ticket-creator-input-section">
                        <label for="parsed-organization">Organization:</label>
                        <select id="parsed-organization">
                            <option value="">Loading organizations...</option>
                        </select>
                    </div>

                    <div class="ticket-creator-input-section">
                        <label for="parsed-user">User:</label>
                        <select id="parsed-user">
                            <option value="">Select an organization first</option>
                        </select>
                    </div>

                    <div class="ticket-creator-input-section" id="computer-selection-section" style="display: none;">
                        <label for="parsed-computer">Computer (Optional):</label>
                        <select id="parsed-computer">
                            <option value="">No computer selected</option>
                        </select>
                        <small style="color: #666; font-size: 11px; display: block; margin-top: 5px;">
                            Computer will be auto-selected if user mentions "their computer" and has exactly one computer.
                        </small>
                    </div>

                    <div class="ticket-creator-input-section">
                        <label for="parsed-subject">Subject:</label>
                        <input type="text" id="parsed-subject" placeholder="Ticket subject">
                    </div>

                    <div class="ticket-creator-input-section">
                        <label for="parsed-issue">Issue Description:</label>
                        <textarea id="parsed-issue" placeholder="Issue description" rows="4"></textarea>
                    </div>

                    <div class="ticket-creator-input-section">
                        <label for="parsed-problem-type">Problem Type:</label>
                        <select id="parsed-problem-type">
                            <option value="Hardware">Hardware</option>
                            <option value="Software">Software</option>
                            <option value="Project / Planned Work">Project / Planned Work</option>
                            <option value="Network / Connectivity">Network / Connectivity</option>
                            <option value="New Device / Deployment">New Device / Deployment</option>
                            <option value="Maintenance / Preventitive">Maintenance / Preventitive</option>
                            <option value="User Account / Access">User Account / Access</option>
                            <option value="Security / Malware">Security / Malware</option>
                            <option value="Internal / MSP Operations">Internal / MSP Operations</option>
                            <option value="Other">Other</option>
                        </select>
                    </div>

                    <div class="ticket-creator-input-section">
                        <label style="display: flex; align-items: center; gap: 8px;">
                            <input type="checkbox" id="parsed-send-email" style="margin: 0;">
                            Send email notification on ticket creation
                        </label>
                    </div>

                    <div style="display: flex; gap: 10px; margin-top: 15px;">
                        <button class="ticket-creator-submit-btn" id="ticket-creator-final-submit-btn" style="flex: 1;">Submit Ticket</button>
                        <button class="ticket-creator-settings-btn" id="ticket-creator-back-btn" style="flex: 1;">Back to Edit</button>
                    </div>
                </div>

                <div id="ticket-creator-result"></div>
            </div>
            <div class="ticket-creator-sidebar-resizer"></div> <!-- New resizer element -->
        `;

        document.body.appendChild(sidebar);

        // Implement resize functionality
        const resizer = sidebar.querySelector('.ticket-creator-sidebar-resizer');
        let isResizing = false;
        let initialX;
        let initialWidth;

        resizer.addEventListener('mousedown', (e) => {
            isResizing = true;
            initialX = e.clientX;
            initialWidth = sidebar.offsetWidth;
            sidebar.classList.add('resizing');
            document.body.style.userSelect = 'none'; // Prevent text selection during drag
            document.body.style.cursor = 'ew-resize'; // Change cursor globally
        });

        document.addEventListener('mousemove', (e) => {
            if (!isResizing) return;

            const dx = e.clientX - initialX;
            let newWidth = initialWidth + dx;

            // Constrain width
            if (newWidth < 350) newWidth = 350;
            if (newWidth > 800) newWidth = 800;

            sidebar.style.width = `${newWidth}px`;
            document.body.style.marginLeft = `${newWidth}px`;
        });

        document.addEventListener('mouseup', () => {
            isResizing = false;
            sidebar.classList.remove('resizing');
            document.body.style.userSelect = '';
            document.body.style.cursor = '';
        });

        // Add debounced input listener to description box for preloading
        const descriptionInput = document.getElementById('ticket-description-input');
        if (descriptionInput) {
            const debouncedPreload = debounce(() => {
                if (!preloadCompleted && !isPreloading) {
                    startPreloadingData();
                }
            }, 500); // 500ms debounce delay
            descriptionInput.addEventListener('input', debouncedPreload);
        }

        // Add close button handler
        document.getElementById('ticket-creator-close-btn').addEventListener('click', () => {
            sidebar.classList.add('minimized');
            document.body.classList.remove('ticket-creator-sidebar-open');
        });

        // Add settings icon button handler
        document.getElementById('ticket-creator-settings-icon-btn').addEventListener('click', () => {
            showConfigModal();
        });

        // Add submit button handler
        document.getElementById('ticket-creator-submit-btn').addEventListener('click', () => {
            handleSendRequest();
        });

        // Add final submit button handler
        document.getElementById('ticket-creator-final-submit-btn').addEventListener('click', () => {
            handleFinalSubmit();
        });

        // Add back button handler
        document.getElementById('ticket-creator-back-btn').addEventListener('click', () => {
            showInitialForm();
        });

        // Load models into dropdown
        fetchOpenRouterModels((models) => {
            const select = document.getElementById('ticket-model-select');
            if (models.length > 0) {
                select.innerHTML = '';

                // Sort models alphabetically by name
                const sortedModels = models.sort((a, b) => a.name.localeCompare(b.name));

                // Add recommended models first (but sorted)
                const recommendedModels = [
                    'anthropic/claude-3-haiku',
                    'anthropic/claude-3.5-sonnet',
                    'openai/gpt-4o',
                    'openai/gpt-4o-mini',
                    'openai/gpt-4-turbo'
                ];

                const recommendedSorted = sortedModels.filter(m => recommendedModels.includes(m.id));
                const otherSorted = sortedModels.filter(m => !recommendedModels.includes(m.id));

                // Add recommended models
                recommendedSorted.forEach(model => {
                    const option = document.createElement('option');
                    option.value = model.id;
                    option.textContent = model.name;
                    select.appendChild(option);
                });

                // Add separator
                const separator = document.createElement('option');
                separator.disabled = true;
                separator.textContent = '────────────';
                select.appendChild(separator);

                // Add all other models (sorted)
                otherSorted.forEach(model => {
                    const option = document.createElement('option');
                    option.value = model.id;
                    option.textContent = model.name;
                    select.appendChild(option);
                });

                // Set default model if configured, otherwise Claude 3.5 Sonnet
                if (DEFAULT_AI_MODEL) {
                    const defaultModel = models.find(m => m.id === DEFAULT_AI_MODEL);
                    if (defaultModel) {
                        select.value = DEFAULT_AI_MODEL;
                    }
                } else {
                    const claudeModel = models.find(m => m.id === 'anthropic/claude-3.5-sonnet');
                    if (claudeModel) {
                        select.value = claudeModel.id;
                    }
                }
            } else {
                select.innerHTML = '<option value="">No models available</option>';
            }
        });

        return sidebar;
    }

    // Handle send request (parse and show review form)
    function handleSendRequest() {
        const description = document.getElementById('ticket-description-input').value.trim();
        const model = document.getElementById('ticket-model-select').value;
        const resultDiv = document.getElementById('ticket-creator-result');
        const submitBtn = document.getElementById('ticket-creator-submit-btn');

        // Validate inputs
        if (!description) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please enter a ticket description.</div>';
            return;
        }

        if (!model) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please select an AI model.</div>';
            return;
        }

        if (!OPENROUTER_API_KEY || !SYNCRO_API_KEY) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please configure your API keys in settings.</div>';
            return;
        }

        // Disable submit button
        submitBtn.disabled = true;
        submitBtn.textContent = 'Processing...';
        resultDiv.innerHTML = '<div class="ticket-creator-result info">Parsing ticket description with AI...</div>';

        // Parse ticket description with AI
        parseTicketDescription(description, model, (error, parsed) => {
            if (error) {
                resultDiv.innerHTML = `<div class="ticket-creator-result error">Error parsing description: ${error}</div>`;
                submitBtn.disabled = false;
                submitBtn.textContent = 'Send';
                return;
            }

            // Debug: log what AI extracted
            console.log('DEBUG: AI parsed result:', parsed);
            console.log('DEBUG: Original description:', description);

            // Store parsed data globally
            parsedTicketData = parsed;

            // Fallback: if AI didn't extract a user, try to extract names from the description
            if (!parsed.user || !parsed.user.trim()) {
                const extractedNames = extractNamesFromDescription(description);
                if (extractedNames.length > 0) {
                    // Use the first extracted name as the user
                    parsed.user = extractedNames[0];
                    parsedTicketData.user = extractedNames[0];
                }
            }

            // Check if we have organization but no user, or user but no organization
            const hasOrganization = parsed.organization && parsed.organization.trim();
            const hasUser = parsed.user && parsed.user.trim();

            if (!hasOrganization && hasUser) {
                // User provided but no organization - search across all customers
                resultDiv.innerHTML = '<div class="ticket-creator-result info">Searching for user across all organizations...</div>';

                findUserAcrossAllCustomers(parsed.user, (error, foundUsers) => {
                    if (error) {
                        resultDiv.innerHTML = `<div class="ticket-creator-result error">Error searching for user: ${error}</div>`;
                        submitBtn.disabled = false;
                        submitBtn.textContent = 'Send';
                        return;
                    }

                    if (foundUsers.length === 0) {
                        // If no contacts found, also search among customer names (some customers might be individuals)
                        resultDiv.innerHTML = '<div class="ticket-creator-result info">No contacts found, searching customers...</div>';

                        // Use preloaded customers for this search
                        if (!preloadCompleted) {
                            resultDiv.innerHTML = `<div class="ticket-creator-result error">Preloading not complete, cannot search customers efficiently. Please wait or try again.</div>`;
                            submitBtn.disabled = false;
                            submitBtn.textContent = 'Send';
                            return;
                        }

                        // Search for the user among preloaded customer names
                        const matchingCustomer = preloadedCustomers.find(c => {
                            const customerName = c.business_name || `${c.firstname} ${c.lastname}`.trim();
                            return customerName && customerName.toLowerCase().includes(parsed.user.toLowerCase());
                        });

                            if (matchingCustomer) {
                                // Found user as a customer - treat them as both the user and organization
                                parsedTicketData.organization = matchingCustomer.business_name || `${matchingCustomer.firstname} ${matchingCustomer.lastname}`;
                                parsedTicketData.user = matchingCustomer.business_name || `${matchingCustomer.firstname} ${matchingCustomer.lastname}`;

                                // Populate dropdowns with this customer
                                populateDropdownsWithCustomer(matchingCustomer, matchingCustomer, parsed);
                            } else {
                                resultDiv.innerHTML = `<div class="ticket-creator-result error">Could not find user "${parsed.user}" in contacts or customers. Please check the name or provide the organization name.</div>`;
                                submitBtn.disabled = false;
                                submitBtn.textContent = 'Send';
                            }
                        return;
                    }

                    if (foundUsers.length === 1) {
                        // Found exactly one user - auto-select organization and user
                        const { user, customer } = foundUsers[0];
                        parsedTicketData.organization = customer.business_name || `${customer.firstname} ${customer.lastname}`;
                        parsedTicketData.user = user.name || `${user.firstname} ${user.lastname}`;

                        // Continue with normal flow using the found customer
                        populateDropdownsWithCustomer(customer, user, parsed);
                    } else {
                        // Multiple users found — display options for user to choose
                        displayUserSelectionOptions(foundUsers, parsed.user);
                        submitBtn.disabled = false;
                        submitBtn.textContent = 'Send';
                        return;
                    }
                });
            } else {
                // Normal flow - use preloaded customers and populate dropdown
                if (!preloadCompleted) {
                    resultDiv.innerHTML = `<div class="ticket-creator-result error">Preloading not complete, cannot populate organizations efficiently. Please wait or try again.</div>`;
                    submitBtn.disabled = false;
                    submitBtn.textContent = 'Send';
                    return;
                }
                populateDropdownsWithCustomers(preloadedCustomers, parsed);
            }
        });
    }

    // Handle final ticket submission
    function handleFinalSubmit() {
        const resultDiv = document.getElementById('ticket-creator-result');
        const submitBtn = document.getElementById('ticket-creator-final-submit-btn');

        if (!parsedTicketData) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">No parsed data available. Please start over.</div>';
            return;
        }

        // Get selected values from dropdowns
        const customerId = document.getElementById('parsed-organization').value;
        const userId = document.getElementById('parsed-user').value;
        const computerId = document.getElementById('parsed-computer').value;
        const subject = document.getElementById('parsed-subject').value.trim();
        const issue = document.getElementById('parsed-issue').value.trim();
        const problemType = document.getElementById('parsed-problem-type').value;
        const sendEmail = document.getElementById('parsed-send-email').checked;

        // Validate required fields
        if (!customerId) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please select an organization.</div>';
            return;
        }
        if (!subject) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please enter a subject.</div>';
            return;
        }
        if (!issue) {
            resultDiv.innerHTML = '<div class="ticket-creator-result error">Please enter an issue description.</div>';
            return;
        }

        // Disable submit button
        submitBtn.disabled = true;
        submitBtn.textContent = 'Creating Ticket...';
        resultDiv.innerHTML = '<div class="ticket-creator-result info">Creating ticket...</div>';

        // Create ticket using selected IDs
        const ticketData = {
            customer_id: customerId,
            subject: subject,
            body: issue,
            problem_type: problemType,
            user_id: userId || null,
            send_email: sendEmail
        };

        // Add asset_ids if a computer was selected
        if (computerId) {
            ticketData.asset_ids = [computerId];
        }

        createSyncroTicket(ticketData, (error, ticket) => {
            if (error) {
                resultDiv.innerHTML = `<div class="ticket-creator-result error">Error creating ticket: ${error}</div>`;
                submitBtn.disabled = false;
                submitBtn.textContent = 'Submit Ticket';
                return;
            }

            // Success - show ticket details
            const domain = USE_SHIELD_DOMAIN ? 'shield.syncromsp.com' : 'syncromsp.com';
            const ticketUrl = `https://${SYNCRO_SUBDOMAIN}.${domain}/tickets/${ticket.id}`;
            resultDiv.innerHTML = `
                <div class="ticket-creator-result success">
                    ✓ Ticket created successfully!<br>
                    <strong>Ticket #${ticket.number}</strong><br>
                    <a href="${ticketUrl}" target="_blank" style="color: #155724; text-decoration: underline;">View Ticket</a>
                </div>
            `;

            // Reset form after successful creation
            setTimeout(() => {
                showInitialForm();
            }, 3000); // Show success message for 3 seconds before resetting
        });
    }

    // Show initial form (hide review section)
    function showInitialForm() {
        const reviewSection = document.getElementById('ticket-review-section');
        if (reviewSection) reviewSection.style.display = 'none';

        const inputSections = document.querySelectorAll('.ticket-creator-input-section');
        inputSections.forEach(section => {
            if (section) section.style.display = 'block';
        });

        const submitBtn = document.getElementById('ticket-creator-submit-btn');
        if (submitBtn) {
            submitBtn.style.display = 'block';
            submitBtn.disabled = false; // Re-enable the button
            submitBtn.textContent = 'Send'; // Reset button text
        }

        const settingsBtn = document.getElementById('ticket-creator-settings-btn');
        if (settingsBtn) settingsBtn.style.display = 'block';

        // Clear the description field
        const descriptionInput = document.getElementById('ticket-description-input');
        if (descriptionInput) descriptionInput.value = '';

        // Reset parsed data
        parsedTicketData = null;

        // Reset result
        const resultDiv = document.getElementById('ticket-creator-result');
        if (resultDiv) resultDiv.innerHTML = '';
    }

    // Setup restore tab
    function setupRestoreTab() {
        let restoreTab = document.getElementById('ticket-creator-restore-tab');
        if (!restoreTab) return;

        restoreTab.addEventListener('click', () => {
            let sidebar = document.getElementById('ticket-creator-sidebar');

            // If sidebar doesn't exist, create it
            if (!sidebar) {
                sidebar = createTicketCreatorSidebar();
            }

            if (sidebar.classList.contains('minimized')) {
                sidebar.classList.remove('minimized');
                document.body.classList.add('ticket-creator-sidebar-open');
            } else {
                sidebar.classList.add('minimized');
                document.body.classList.remove('ticket-creator-sidebar-open');
            }
        });
    }

    // Initialize when page loads
    function initialize() {
        // Prompt for API keys if not set
        promptForApiKeys();

        // Create shared styles
        createSharedStyles();

        // Create restore tab
        let restoreTab = document.getElementById('ticket-creator-restore-tab');
        if (!restoreTab) {
            restoreTab = document.createElement('button');
            restoreTab.id = 'ticket-creator-restore-tab';
            restoreTab.className = 'ticket-creator-restore-tab';
            restoreTab.textContent = '🎫 Create Ticket';
            document.body.appendChild(restoreTab);
        }

        // Setup tab click handler
        setupRestoreTab();
    }

    // Wait for page to load
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize);
    } else {
        initialize();
    }

})();


