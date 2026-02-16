(function (window, undefined) {
    const PLUGIN_VERSION = '4.3';
    const API_URL = 'http://24.227.108.165:9000/editor-chatbot';
    const IMAGE_API_URL = 'http://24.227.108.165:9000/api/image'; // Image generation API
    const MAX_MESSAGES = 5;
    const MAX_IMAGE_GENERATIONS = 2; // Maximum 2 image generations
    const MAX_CONTEXT_CHARS = 30000; // 30k character limit

    // Initialization guard
    let isInitialized = false;

    let messageCount = 0;
    let currentContext = {};
    let documentHeaders = [];
    let rfpTitle = '';
    let selectedText = '';
    let totalContextChars = 0;

    // Image generator state
    let currentImageCallId = null; // Stores image_call_id for modifications
    let currentConversationId = null; // Stores conversation_id
    let imageGenerationCount = 0;

    // DOM Elements
    let chatContainer, messageInput, sendBtn, attachBtn, clearContextBtn;
    let contextContainer, contextContent, headerSuggestions, chatCount;
    let chatTab, imageTab;

    // Current mode
    let currentMode = 'chat';

    window.Asc.plugin.init = function () {
        // Prevent multiple initializations
        if (isInitialized) {
            console.warn('%c[INIT] Plugin already initialized, skipping...', 'color: #ff9800; font-weight: bold;');
            console.warn('%c[INIT] This is initialization attempt #' + (window._initAttempts || 1), 'color: #ff9800;');
            return;
        }

        // Track initialization attempts
        window._initAttempts = (window._initAttempts || 0) + 1;

        console.log('%c=== AI CHATBOT PLUGIN ===', 'color: #667eea; font-weight: bold; font-size: 14px;');
        console.log('%c[INIT] Attempt #' + window._initAttempts, 'color: #667eea; font-weight: bold;');
        console.log('%cVersion: ' + PLUGIN_VERSION, 'color: #764ba2; font-weight: bold;');
        console.log('%cTimestamp: ' + new Date().toISOString(), 'color: #666;');
        console.log('%cScript URL: ' + document.querySelector('script[src*="pluginCode.js"]')?.src, 'color: #666;');
        console.log('%cAPI URL: ' + API_URL, 'color: #666;');
        console.log('%cImage API URL: ' + IMAGE_API_URL, 'color: #666;');
        console.log('%cMax Messages: ' + MAX_MESSAGES, 'color: #666;');
        console.log('%cMax Context Chars: ' + MAX_CONTEXT_CHARS.toLocaleString(), 'color: #666;');
        console.log('%c================================', 'color: #667eea;');

        // Detect and apply editor theme
        detectAndApplyTheme();

        initElements();
        getDocumentTitle();

        // Delay header extraction to ensure document API is ready
        setTimeout(function() {
            extractHeaders();
        }, 500);

        setupEventListeners();

        // Set initial welcome message for chat mode
        setWelcomeMessage('chat');

        // Check if opened in image mode
        checkPluginMode();

        // Mark as initialized
        isInitialized = true;
        console.log('%c[INIT] Initialization complete!', 'color: #4caf50; font-weight: bold;');
    };

    function setWelcomeMessage(mode) {
        if (mode === 'chat') {
            chatContainer.innerHTML = '<div class="welcome-message">' +
                '<p>👋 Welcome! I\'m your AI assistant for collaborative document editing.</p>' +
                '<p>You can:</p>' +
                '<ul>' +
                '<li>Ask questions about your document</li>' +
                '<li>Select text and attach it as context</li>' +
                '<li>Use @ to reference headers</li>' +
                '</ul>' +
                '</div>';
        } else {
            chatContainer.innerHTML = '<div class="welcome-message">' +
                '<p>🖼️ Image Generator Mode</p>' +
                '<p>Describe the image you want to create and I\'ll generate it for you.</p>' +
                '<p><strong>Note:</strong> You can generate up to ' + MAX_IMAGE_GENERATIONS + ' images per session.</p>' +
                '</div>';
        }
    }

    function detectAndApplyTheme() {
        // Check URL parameters for theme
        const urlParams = new URLSearchParams(window.location.search);
        const theme = urlParams.get('theme') || 'theme-light';

        console.log('%c[THEME] Editor theme detected:', 'color: #ff9800; font-weight: bold;', theme);

        // Apply theme to plugin container
        const body = document.body;

        if (theme === 'theme-dark' || theme.includes('dark')) {
            console.log('%c[THEME] Applying dark theme styles...', 'color: #ff9800;');
            body.classList.add('dark-theme');
            body.style.background = '#1e1e1e';
            body.style.color = '#e0e0e0';
        } else if (theme.includes('contrast-dark')) {
            console.log('%c[THEME] Applying high contrast dark theme styles...', 'color: #ff9800;');
            body.classList.add('dark-theme', 'high-contrast');
            body.style.background = '#000000';
            body.style.color = '#ffffff';
        } else {
            console.log('%c[THEME] Applying light theme styles...', 'color: #ff9800;');
            body.classList.add('light-theme');
            body.style.background = '#ffffff';
            body.style.color = '#333333';
        }
    }

    function checkPluginMode() {
        // Check URL parameters to determine mode
        const urlParams = new URLSearchParams(window.location.search);
        const mode = urlParams.get('mode');

        console.log('%c[MODE] Plugin mode:', 'color: #9c27b0; font-weight: bold;', mode || 'chat (default)');

        if (mode === 'image') {
            // Show a welcome message for image generator
            showMessage('system', 'Image Generator mode. Describe the image you want to create.');
        }
    }

    function initElements() {
        chatContainer = document.getElementById('chatContainer');
        messageInput = document.getElementById('messageInput');
        sendBtn = document.getElementById('sendBtn');
        attachBtn = document.getElementById('attachBtn');
        clearContextBtn = document.getElementById('clearContext');
        contextContainer = document.getElementById('contextContainer');
        contextContent = document.getElementById('contextContent');
        headerSuggestions = document.getElementById('headerSuggestions');
        chatCount = document.getElementById('chatCount');
        chatTab = document.getElementById('chatTab');
        imageTab = document.getElementById('imageTab');
    }

    function setupEventListeners() {
        sendBtn.addEventListener('click', handleSendMessage);
        attachBtn.addEventListener('click', attachSelectedText);
        clearContextBtn.addEventListener('click', clearContext);

        messageInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleSendMessage();
            }
        });

        messageInput.addEventListener('input', handleInputChange);

        // Tab switching
        chatTab.addEventListener('click', function() {
            switchMode('chat');
        });

        imageTab.addEventListener('click', function() {
            switchMode('image');
        });
    }

    function switchMode(mode) {
        currentMode = mode;

        // Clear context
        currentContext = {};
        totalContextChars = 0;
        contextContainer.style.display = 'none';

        // Hide header suggestions when switching modes
        headerSuggestions.style.display = 'none';

        if (mode === 'chat') {
            chatTab.classList.add('active');
            imageTab.classList.remove('active');

            // Show attach button and context features
            attachBtn.style.display = 'inline-block';
            messageInput.placeholder = 'Type your question or use @ to reference headers...';

            // Clear image generation state when switching to chat
            currentImageCallId = null;
            currentConversationId = null;

            // Update count display for chat
            updateChatCount();
        } else {
            chatTab.classList.remove('active');
            imageTab.classList.add('active');

            // Hide attach button and context features
            attachBtn.style.display = 'none';
            messageInput.placeholder = 'Describe the image you want to generate...';

            // Update count display for images
            updateImageCount();
        }

        // Set welcome message and clear chat
        setWelcomeMessage(mode);

        console.log('%c[MODE] Switched to:', 'color: #9c27b0; font-weight: bold;', mode);
    }

    function getDocumentTitle() {
        // Use callCommand to get the actual document name from the editor
        window.Asc.plugin.callCommand(function() {
            return Api.GetFullName();
        }, false, true, function(result) {
            if (result && result.trim() !== '' && result !== 'Untitled Document') {
                // Remove file extension if present
                rfpTitle = result.replace(/\.(docx|doc)$/i, '');
                console.log('%c[TITLE] Document title:', 'color: #00bcd4;', rfpTitle);
            } else {
                // Fallback: Try GetFileInfo
                window.Asc.plugin.executeMethod("GetFileInfo", [], function(info) {
                    let titleFound = false;

                    if (info) {
                        // Try Title property
                        if (info.Title && info.Title !== 'Untitled Document' && info.Title.trim() !== '') {
                            rfpTitle = info.Title;
                            titleFound = true;
                        }
                        // Try FileName property (without extension)
                        else if (info.FileName) {
                            rfpTitle = info.FileName.replace(/\.(docx|doc)$/i, '');
                            titleFound = true;
                        }
                    }

                    // Update display
                    if (!titleFound) {
                        // Default fallback
                        rfpTitle = 'RFP Document';
                    }
                    console.log('%c[TITLE] Document title:', 'color: #00bcd4;', rfpTitle);
                });
            }
        });
    }

    function extractHeaders() {
        window.Asc.plugin.callCommand(function() {
            var oDocument = Api.GetDocument();
            var headingsList = [];
            var totalElements = oDocument.GetElementsCount();

            for (var i = 0; i < totalElements; i++) {
                var oElement = oDocument.GetElement(i);

                if (oElement.GetClassType() === "paragraph") {
                    var oStyle = oElement.GetStyle();

                    if (oStyle !== null && oStyle !== undefined) {
                        var styleName = oStyle.GetName();

                        if (styleName.indexOf("Heading") !== -1) {
                            var text = oElement.GetText().replace(/\r/g, "");

                            if (text && text.trim()) {
                                headingsList.push({
                                    text: text.trim(),
                                    styleName: styleName,
                                    index: i
                                });
                            }
                        }
                    }
                }
            }
            return JSON.stringify(headingsList);
        }, false, true, function(result) {
            try {
                documentHeaders = [];
                var headings = null;

                console.log('extractHeaders callback - result type:', typeof result);
                console.log('extractHeaders callback - result:', result);

                // Handle both array and JSON string responses
                if (result) {
                    if (Array.isArray(result)) {
                        // Direct array response
                        console.log('Result is array with length:', result.length);
                        headings = result;
                    } else if (typeof result === 'string') {
                        var trimmedResult = result.trim();
                        if (trimmedResult.startsWith('[') || trimmedResult.startsWith('{')) {
                            // JSON string response
                            console.log('Parsing JSON string...');
                            headings = JSON.parse(result);
                        } else {
                            console.warn('⚠ Received non-JSON string (possibly filename). Document API may not be ready yet.');
                            console.log('Will retry in 1 second...');
                            // Retry after 1 second
                            setTimeout(function() {
                                console.log('Retrying header extraction...');
                                extractHeaders();
                            }, 1000);
                            return;
                        }
                    }

                    if (headings && Array.isArray(headings) && headings.length > 0) {
                        headings.forEach(function(heading) {
                            var level = extractLevelFromStyle(heading.styleName);
                            documentHeaders.push({
                                text: heading.text,
                                level: heading.styleName,
                                index: heading.index,
                                outlineLevel: level
                            });
                        });
                        console.log('✓ Extracted headers:', documentHeaders.length);
                        documentHeaders.forEach(function(h) {
                            console.log('  -', h.level, ':', h.text);
                        });
                    } else {
                        console.log('No headings found in document');
                    }
                }
            } catch (error) {
                console.error('Error processing headers:', error);
                console.log('Raw result:', result);
            }
        });
    }

    function extractLevelFromStyle(styleName) {
        var match = styleName.match(/[1-9]/);
        if (match) {
            return parseInt(match[0]) - 1;
        }
        return 0; // Default to level 0 if no number found
    }

    function handleInputChange() {
        // Only handle @ feature in chat mode
        if (currentMode !== 'chat') {
            return;
        }

        const text = messageInput.value;
        const lastAtSymbol = text.lastIndexOf('@');

        if (lastAtSymbol !== -1 && lastAtSymbol === text.length - 1) {
            showHeaderSuggestions();
        } else if (lastAtSymbol !== -1) {
            const searchTerm = text.substring(lastAtSymbol + 1).toLowerCase();
            filterHeaderSuggestions(searchTerm);
        } else {
            hideHeaderSuggestions();
        }
    }

    function showHeaderSuggestions() {
        if (documentHeaders.length === 0) {
            headerSuggestions.innerHTML = '<div class="suggestion-item disabled">No headers found in document</div>';
            headerSuggestions.style.display = 'block';
            return;
        }

        headerSuggestions.innerHTML = '';
        documentHeaders.slice(0, 10).forEach(function(header) {
            const item = document.createElement('div');
            item.className = 'suggestion-item';
            item.textContent = header.text;
            item.addEventListener('click', function() {
                selectHeader(header);
            });
            headerSuggestions.appendChild(item);
        });
        headerSuggestions.style.display = 'block';
    }

    function filterHeaderSuggestions(searchTerm) {
        const filtered = documentHeaders.filter(function(header) {
            return header.text.toLowerCase().includes(searchTerm);
        });

        if (filtered.length === 0) {
            hideHeaderSuggestions();
            return;
        }

        headerSuggestions.innerHTML = '';
        filtered.slice(0, 10).forEach(function(header) {
            const item = document.createElement('div');
            item.className = 'suggestion-item';
            item.textContent = header.text;
            item.addEventListener('click', function() {
                selectHeader(header);
            });
            headerSuggestions.appendChild(item);
        });
        headerSuggestions.style.display = 'block';
    }

    function hideHeaderSuggestions() {
        headerSuggestions.style.display = 'none';
    }

    function selectHeader(header) {
        console.log('=== SELECTING HEADER ===');
        console.log('Header:', header.text);
        console.log('Level:', header.level);

        // Get header content
        getHeaderContent(header, function(content) {
            console.log('Retrieved content length:', content.length);
            console.log('Content preview:', content.substring(0, 100) + '...');

            const contextKey = Object.keys(currentContext).length + 1;

            if (contextKey > 2) {
                console.warn('Maximum contexts reached (2)');
                showMessage('system', 'Maximum 2 contexts can be attached at once.');
                return;
            }

            // Check character limit
            const contentLength = content.length;
            const remainingChars = MAX_CONTEXT_CHARS - totalContextChars;

            console.log('Current total chars:', totalContextChars);
            console.log('Content length:', contentLength);
            console.log('Remaining chars:', remainingChars);

            if (contentLength > MAX_CONTEXT_CHARS) {
                console.error('Content exceeds maximum limit:', contentLength, '>', MAX_CONTEXT_CHARS);
                showMessage('system', 'Context too long! This section has ' + formatNumber(contentLength) + ' characters. Maximum allowed is ' + formatNumber(MAX_CONTEXT_CHARS) + ' characters.');
                hideHeaderSuggestions();
                return;
            }

            if (contentLength > remainingChars) {
                console.error('Content exceeds remaining limit:', contentLength, '>', remainingChars);
                showMessage('system', 'Context too long! You have ' + formatNumber(totalContextChars) + ' characters already. This section (' + formatNumber(contentLength) + ' chars) exceeds the remaining limit of ' + formatNumber(remainingChars) + ' characters.');
                hideHeaderSuggestions();
                return;
            }

            currentContext[contextKey] = {
                content: content,
                isHeader: true,
                headerText: header.text,
                charCount: contentLength
            };

            totalContextChars += contentLength;

            console.log('✓ Header attached successfully');
            console.log('Context key:', contextKey);
            console.log('New total chars:', totalContextChars);

            // Highlight the header content in the document
            highlightText(header.text);

            // Remove @ and search term from input
            const text = messageInput.value;
            const lastAtSymbol = text.lastIndexOf('@');
            messageInput.value = text.substring(0, lastAtSymbol);

            hideHeaderSuggestions();
            updateContextDisplay();
            console.log('=== HEADER SELECTION COMPLETE ===');
        });
    }

    function getHeaderContent(header, callback) {
        console.log('=== GET HEADER CONTENT ===');
        console.log('Searching for header:', header.text);
        console.log('Header level:', header.level);

        // Escape values for safe embedding in code string
        var searchPhrase = JSON.stringify(header.text);
        var targetLevel = JSON.stringify(header.level);

        // Create function dynamically with embedded values
        var func = new Function('return function() {' +
            'var searchPhrase = ' + searchPhrase + ';' +
            'var targetLevel = ' + targetLevel + ';' +
            'var oDocument = Api.GetDocument();' +
            'var totalElements = oDocument.GetElementsCount();' +
            'var foundText = [];' +
            'var isCollecting = false;' +
            'var targetHeadingNumber = null;' +
            'var levelMatch = targetLevel.match(/[1-9]/);' +
            'if (levelMatch) { targetHeadingNumber = parseInt(levelMatch[0]); }' +
            'console.log("Target heading number:", targetHeadingNumber);' +
            'console.log("Total elements in document:", totalElements);' +
            'for (var i = 0; i < totalElements; i++) {' +
                'var oElement = oDocument.GetElement(i);' +
                'if (oElement.GetClassType() === "paragraph") {' +
                    'var oStyle = oElement.GetStyle();' +
                    'var text = oElement.GetText().replace(/\\r/g, "").trim();' +
                    'var isHeading = false;' +
                    'var currentStyleName = "";' +
                    'var currentHeadingNumber = null;' +
                    'if (oStyle !== null && oStyle !== undefined) {' +
                        'currentStyleName = oStyle.GetName();' +
                        'if (currentStyleName.indexOf("Heading") !== -1) {' +
                            'isHeading = true;' +
                            'var headingMatch = currentStyleName.match(/[1-9]/);' +
                            'if (headingMatch) { currentHeadingNumber = parseInt(headingMatch[0]); }' +
                        '}' +
                    '}' +
                    'if (isHeading) {' +
                        'console.log("Found heading at index", i, ":", currentStyleName, "-", text);' +
                        'if (text.toLowerCase().indexOf(searchPhrase.toLowerCase()) !== -1) {' +
                            'console.log("✓ Match found! Starting collection...");' +
                            'isCollecting = true;' +
                            'foundText.push(text);' +
                        '} else if (isCollecting) {' +
                            'if (currentHeadingNumber !== null && targetHeadingNumber !== null) {' +
                                'if (currentHeadingNumber <= targetHeadingNumber) {' +
                                    'console.log("Stopping collection - found same/higher level heading:", currentStyleName);' +
                                    'isCollecting = false;' +
                                '} else {' +
                                    'console.log("Sub-heading found, continuing collection:", currentStyleName);' +
                                    'foundText.push(text);' +
                                '}' +
                            '} else {' +
                                'console.log("Stopping collection - found another heading");' +
                                'isCollecting = false;' +
                            '}' +
                        '}' +
                    '} else {' +
                        'if (isCollecting && text !== "") {' +
                            'console.log("Collecting text at index", i, ":", text.substring(0, 50) + "...");' +
                            'foundText.push(text);' +
                        '}' +
                    '}' +
                '}' +
            '}' +
            'console.log("Collection complete. Total items collected:", foundText.length);' +
            'return foundText.join("\\n\\n");' +
        '}')();

        window.Asc.plugin.callCommand(func, false, true, function(result) {
            console.log('Header content retrieved. Length:', result ? result.length : 0);
            console.log('=== END GET HEADER CONTENT ===');
            callback(result || '');
        });
    }

    function attachSelectedText() {
        console.log('=== ATTACHING SELECTED TEXT ===');
        window.Asc.plugin.executeMethod("GetSelectedText", [true], function(text) {
            if (!text || text.trim() === '') {
                console.warn('No text selected');
                showMessage('system', 'Please select some text first.');
                return;
            }

            console.log('Selected text length:', text.length);
            console.log('Text preview:', text.substring(0, 100) + '...');

            const contextKey = Object.keys(currentContext).length + 1;

            if (contextKey > 2) {
                console.warn('Maximum contexts reached (2)');
                showMessage('system', 'Maximum 2 contexts can be attached at once.');
                return;
            }

            // Check character limit
            const contentLength = text.length;
            const remainingChars = MAX_CONTEXT_CHARS - totalContextChars;

            console.log('Current total chars:', totalContextChars);
            console.log('Content length:', contentLength);
            console.log('Remaining chars:', remainingChars);

            if (contentLength > MAX_CONTEXT_CHARS) {
                console.error('Content exceeds maximum limit:', contentLength, '>', MAX_CONTEXT_CHARS);
                showMessage('system', 'Context too long! Selected text has ' + formatNumber(contentLength) + ' characters. Maximum allowed is ' + formatNumber(MAX_CONTEXT_CHARS) + ' characters.');
                return;
            }

            if (contentLength > remainingChars) {
                console.error('Content exceeds remaining limit:', contentLength, '>', remainingChars);
                showMessage('system', 'Context too long! You have ' + formatNumber(totalContextChars) + ' characters already. Selected text (' + formatNumber(contentLength) + ' chars) exceeds the remaining limit of ' + formatNumber(remainingChars) + ' characters.');
                return;
            }

            currentContext[contextKey] = {
                content: text,
                isHeader: false,
                charCount: contentLength
            };
            selectedText = text;
            totalContextChars += contentLength;

            console.log('✓ Text attached successfully');
            console.log('Context key:', contextKey);
            console.log('New total chars:', totalContextChars);

            // Highlight the selected text in the document
            highlightSelectedText();

            updateContextDisplay();
            console.log('=== TEXT ATTACHMENT COMPLETE ===');
        });
    }

    function highlightText(searchText) {
        // Use ONLYOFFICE API to highlight text
        window.Asc.plugin.executeMethod("SearchAndReplace", [{
            searchString: searchText,
            replaceString: searchText,
            matchCase: false,
            highlightColor: [255, 255, 153] // Light yellow highlight
        }]);
    }

    function highlightSelectedText() {
        // Apply highlight to current selection
        window.Asc.plugin.executeMethod("SetTextProperties", [{
            HighlightColor: [255, 255, 153] // Light yellow highlight
        }]);
    }

    function updateContextDisplay() {
        if (Object.keys(currentContext).length === 0) {
            contextContainer.style.display = 'none';
            return;
        }

        contextContent.innerHTML = '';
        Object.keys(currentContext).forEach(function(key) {
            const contextItem = document.createElement('div');
            contextItem.className = 'context-item';
            contextItem.style.cssText = 'display: flex; justify-content: space-between; align-items: center; padding: 6px 8px; margin-bottom: 6px;';

            const ctx = currentContext[key];
            const content = typeof ctx === 'string' ? ctx : ctx.content;
            const isHeader = ctx.isHeader || false;
            const headerText = ctx.headerText || '';
            const charCount = ctx.charCount || content.length;

            const preview = content.substring(0, 40) +
                           (content.length > 40 ? '...' : '');

            const label = isHeader ? headerText : preview;

            // Create content container
            const contentDiv = document.createElement('div');
            contentDiv.style.cssText = 'flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; margin-right: 8px;';
            contentDiv.innerHTML = '<span style="font-weight: 500;">' + escapeHtml(label) + '</span> <span style="color: #666; font-size: 11px;">(' + formatNumber(charCount) + ')</span>';

            // Create remove button
            const removeBtn = document.createElement('button');
            removeBtn.textContent = '✕';
            removeBtn.className = 'context-remove-btn';
            removeBtn.style.cssText = 'background: none; border: none; color: #d32f2f; cursor: pointer; font-size: 16px; padding: 0; width: 20px; height: 20px; display: flex; align-items: center; justify-content: center; border-radius: 50%; transition: background 0.2s;';
            removeBtn.onclick = function() {
                removeContext(key);
            };

            // Hover effect
            removeBtn.onmouseenter = function() {
                removeBtn.style.background = 'rgba(211, 47, 47, 0.1)';
            };
            removeBtn.onmouseleave = function() {
                removeBtn.style.background = 'none';
            };

            contextItem.appendChild(contentDiv);
            contextItem.appendChild(removeBtn);
            contextContent.appendChild(contextItem);
        });

        // Add total character count display
        const totalDisplay = document.createElement('div');
        totalDisplay.style.cssText = 'margin-top: 8px; padding-top: 8px; border-top: 1px solid #90caf9; font-size: 12px; font-weight: 600; color: #1976d2;';
        totalDisplay.textContent = 'Total: ' + formatNumber(totalContextChars) + ' / ' + formatNumber(MAX_CONTEXT_CHARS) + ' characters';
        contextContent.appendChild(totalDisplay);

        contextContainer.style.display = 'block';
    }

    function removeContext(key) {
        console.log('=== REMOVING CONTEXT ===');
        console.log('Removing context key:', key);

        if (currentContext[key]) {
            const charCount = currentContext[key].charCount || currentContext[key].content.length;
            totalContextChars -= charCount;
            delete currentContext[key];

            console.log('Context removed. Remaining chars:', totalContextChars);
            updateContextDisplay();
        }
    }

    function clearContext() {
        console.log('=== CLEARING CONTEXT ===');
        console.log('Previous context count:', Object.keys(currentContext).length);
        console.log('Previous total chars:', totalContextChars);
        currentContext = {};
        selectedText = '';
        totalContextChars = 0;
        contextContainer.style.display = 'none';
        console.log('✓ Context cleared');
    }

    function handleSendMessage() {
        // Prevent double submission
        if (sendBtn.disabled) {
            return;
        }

        const userInput = messageInput.value.trim();
        if (!userInput) {
            showMessage('system', currentMode === 'image' ? 'Please enter an image description.' : 'Please enter a question.');
            return;
        }

        // Route based on current mode
        if (currentMode === 'image') {
            handleImageGeneration(userInput);
        } else {
            handleChatMessage(userInput);
        }
    }

    function handleChatMessage(question) {
        if (messageCount >= MAX_MESSAGES) {
            showMessage('system', 'Maximum message limit reached (5 messages). Please restart the plugin.');
            return;
        }

        // Disable send button immediately to prevent double submission
        sendBtn.disabled = true;
        sendBtn.textContent = 'Sending...';

        // Show user message
        showMessage('user', question);

        // Clear input
        messageInput.value = '';

        // Prepare API payload
        const payload = {
            rfp_title: rfpTitle,
            question: question
        };

        if (Object.keys(currentContext).length > 0) {
            // Convert context to expected format
            const contextData = {};
            Object.keys(currentContext).forEach(function(key) {
                const ctx = currentContext[key];
                contextData[key] = typeof ctx === 'string' ? ctx : ctx.content;
            });
            payload.context = contextData;
        }

        // Call chatbot API
        callAPI(payload);
    }

    function handleImageGeneration(description) {
        if (imageGenerationCount >= MAX_IMAGE_GENERATIONS) {
            showMessage('system', 'Maximum image generation limit reached (' + MAX_IMAGE_GENERATIONS + ' images). Please restart the plugin.');
            return;
        }

        // Disable send button
        sendBtn.disabled = true;
        sendBtn.textContent = 'Generating...';

        // Show user message
        showMessage('user', description);

        // Clear input
        messageInput.value = '';

        // Call image generation API
        callImageAPI(description);
    }

    function callImageAPI(description) {
        console.log('=== IMAGE API CALL START ===');
        console.log('API URL:', IMAGE_API_URL);
        console.log('Description:', description);
        console.log('Current image_call_id:', currentImageCallId || 'none (new image)');
        console.log('Current conversation_id:', currentConversationId || 'none');

        // Build payload with image_call_id and conversation_id
        const payload = {
            prompt: description,
            image_call_id: currentImageCallId || null,
            conversation_id: currentConversationId || null
        };

        console.log('Payload:', JSON.stringify(payload, null, 2));

        const abortController = new AbortController();
        const timeoutId = setTimeout(function() {
            abortController.abort();
        }, 60000); // 60 second timeout for image generation

        fetch(IMAGE_API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload),
            signal: abortController.signal
        })
        .then(function(response) {
            clearTimeout(timeoutId);
            console.log('Response status:', response.status);
            return response.json();
        })
        .then(function(data) {
            console.log('Response data received:', JSON.stringify(data, null, 2));
            console.log('=== IMAGE API CALL END ===');
            handleImageAPIResponse(data);
        })
        .catch(function(error) {
            clearTimeout(timeoutId);
            console.error('Image API call error:', error);
            console.log('=== IMAGE API CALL FAILED ===');

            let errorMessage = 'Image generation failed';
            if (error.name === 'AbortError') {
                errorMessage = 'Image generation timeout - please try again';
            }

            showMessage('bot', errorMessage);
            sendBtn.disabled = false;
            sendBtn.textContent = 'Send';
        });
    }

    function handleImageAPIResponse(data) {
        console.log('%c[IMAGE] Processing API response...', 'color: #9c27b0; font-weight: bold;');

        if (data.success && data.image_base64) {
            // Store image_call_id and conversation_id for future modifications
            if (data.image_call_id) {
                currentImageCallId = data.image_call_id;
                console.log('[IMAGE] Stored image_call_id:', currentImageCallId);
            }

            if (data.conversation_id) {
                currentConversationId = data.conversation_id;
                console.log('[IMAGE] Stored conversation_id:', currentConversationId);
            }

            // Increment count only for new images (not modifications)
            if (!currentImageCallId || imageGenerationCount === 0) {
                imageGenerationCount++;
                updateImageCount();
            }

            showMessage('bot', data.message || 'Image generated successfully! Inserting into document...');

            // Insert image into document using base64 data
            insertImageIntoDocument(data.image_base64);
        } else {
            showMessage('system', '⚠ Image generation failed: ' + (data.message || data.error || 'Unknown error'));
        }

        sendBtn.disabled = false;
        sendBtn.textContent = 'Send';
    }

    function insertImageIntoDocument(imageBase64) {
        console.log('%c[IMAGE] Inserting image into document...', 'color: #9c27b0; font-weight: bold;');

        if (!imageBase64) {
            console.error('[IMAGE] No image data provided');
            showMessage('system', '⚠ Error: No image data received');
            return;
        }

        // Convert base64 to data URL if not already in that format
        let imageDataUrl = imageBase64;
        if (!imageBase64.startsWith('data:')) {
            imageDataUrl = 'data:image/png;base64,' + imageBase64;
        }

        console.log('[IMAGE] Image data URL length:', imageDataUrl.length);
        console.log('[IMAGE] Using InsertImage method...');

        // Try using InsertImage method with proper parameters
        try {
            window.Asc.plugin.executeMethod('InsertImage', [{
                'c': 'add',
                'images': [{
                    'src': imageDataUrl,
                    'width': 400,
                    'height': 300
                }]
            }], function(result) {
                console.log('[IMAGE] InsertImage result:', result);
                if (result !== false && result !== null) {
                    console.log('%c[IMAGE] Image inserted successfully', 'color: #4caf50; font-weight: bold;');
                    showMessage('system', '✓ Image inserted into document');
                } else {
                    console.warn('[IMAGE] InsertImage returned false/null, trying PasteHtml...');
                    // Fallback: Try PasteHtml with img tag
                    const imgHtml = '<img src="' + imageDataUrl + '" width="400" height="300" />';
                    window.Asc.plugin.executeMethod('PasteHtml', [imgHtml], function(pasteResult) {
                        console.log('[IMAGE] PasteHtml result:', pasteResult);
                        if (pasteResult !== false) {
                            console.log('%c[IMAGE] Image inserted via PasteHtml', 'color: #4caf50; font-weight: bold;');
                            showMessage('system', '✓ Image inserted into document');
                        } else {
                            console.error('[IMAGE] Both methods failed');
                            showMessage('system', '⚠ Failed to insert image. Try copying the image manually.');
                        }
                    });
                }
            });
        } catch (error) {
            console.error('[IMAGE] Error during insertion:', error);
            showMessage('system', '⚠ Error inserting image: ' + error.message);
        }
    }

    function callAPI(payload) {
        console.log('=== CHATBOT API CALL START ===');
        console.log('API URL:', API_URL);
        console.log('Method: POST');
        console.log('Payload:', JSON.stringify(payload, null, 2));

        // Set 60 second timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(function() {
            controller.abort();
        }, 60000);

        fetch(API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload),
            signal: controller.signal
        })
        .then(function(response) {
            clearTimeout(timeoutId);
            console.log('Response status:', response.status);
            console.log('Response headers:', response.headers);

            if (!response.ok) {
                throw new Error('Server error occurred');
            }
            return response.json();
        })
        .then(function(data) {
            console.log('Response data received:', JSON.stringify(data, null, 2));
            console.log('=== CHATBOT API CALL END ===');
            handleAPIResponse(data);
        })
        .catch(function(error) {
            clearTimeout(timeoutId);
            console.error('API call error:', error);
            console.log('=== CHATBOT API CALL FAILED ===');

            // Show generic error message
            let errorMessage = 'Server error occurred';
            if (error.name === 'AbortError') {
                errorMessage = 'Request timeout - please try again';
            }

            showMessage('bot', errorMessage);

            // Re-enable send button
            sendBtn.disabled = false;
            sendBtn.textContent = 'Send';
        });
    }

    function handleAPIResponse(data) {
        console.log('=== HANDLING API RESPONSE ===');
        console.log('Response status:', data.status);

        if (data.status !== 'success') {
            console.error('API returned error status');
            showMessage('system', 'API Error: ' + (data.message || 'Unknown error'));
            sendBtn.disabled = false;
            sendBtn.textContent = 'Send';
            return;
        }

        const response = data.response;
        console.log('Bot response:', response.bot_response);
        console.log('Replacement text:', response.replacement_text);

        // Show bot response
        showMessage('bot', response.bot_response);

        // Handle replacements
        if (response.replacement_text && Object.keys(response.replacement_text).length > 0) {
            console.log('Processing replacements...');
            handleReplacements(response.replacement_text);
        } else {
            console.log('No replacements to apply');
        }

        // Increment message count
        messageCount++;
        console.log('Message count:', messageCount);
        updateChatCount();

        // Clear context after successful send
        clearContext();

        // Re-enable send button
        sendBtn.disabled = false;
        sendBtn.textContent = 'Send';
        console.log('=== API RESPONSE HANDLED ===');
    }

    function handleReplacements(replacementText) {
        console.log('=== HANDLING REPLACEMENTS ===');
        let replacementCount = 0;
        let failedReplacements = [];

        Object.keys(replacementText).forEach(function(contextKey) {
            const replacements = replacementText[contextKey];

            // Skip if replacements is not an array or is empty
            if (!Array.isArray(replacements) || replacements.length === 0) {
                console.log('Context key:', contextKey, '- No replacements (empty or invalid)');
                return;
            }

            console.log('Context key:', contextKey, 'Replacements:', replacements.length);

            replacements.forEach(function(replacement, index) {
                // Validate replacement object
                if (!replacement || !replacement.target_text || !replacement.new_text) {
                    console.warn('Skipping invalid replacement at index', index, ':', replacement);
                    return;
                }

                const targetText = replacement.target_text;
                const newText = replacement.new_text;

                console.log('Replacement', index + 1, ':');
                console.log('  Target:', targetText.substring(0, 50) + '...');
                console.log('  New:', newText.substring(0, 50) + '...');

                try {
                    searchAndReplace(targetText, newText);
                    replacementCount++;
                } catch (error) {
                    console.error('Replacement failed:', error);
                    failedReplacements.push({
                        target: targetText.substring(0, 50) + '...',
                        error: error.message
                    });
                }
            });
        });

        console.log('Total replacements applied:', replacementCount);

        // Show success message in chat only if replacements were actually made
        if (replacementCount > 0) {
            showMessage('system', '✓ ' + replacementCount + ' text replacement(s) applied to document.');
        }

        // Show failed replacements in chat if any
        if (failedReplacements.length > 0) {
            const failedMsg = '⚠ ' + failedReplacements.length + ' replacement(s) failed. Text not found in document.';
            showMessage('system', failedMsg);
        }

        console.log('=== REPLACEMENTS COMPLETE ===');
    }

    function calculateSimilarity(str1, str2) {
        // Simple Levenshtein distance for fuzzy matching
        const len1 = str1.length;
        const len2 = str2.length;
        const matrix = [];

        for (let i = 0; i <= len1; i++) {
            matrix[i] = [i];
        }
        for (let j = 0; j <= len2; j++) {
            matrix[0][j] = j;
        }

        for (let i = 1; i <= len1; i++) {
            for (let j = 1; j <= len2; j++) {
                if (str1.charAt(i - 1).toLowerCase() === str2.charAt(j - 1).toLowerCase()) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }

        const distance = matrix[len1][len2];
        const maxLen = Math.max(len1, len2);
        const similarity = ((maxLen - distance) / maxLen) * 100;
        return similarity;
    }

    function searchAndReplace(searchText, replaceText) {
        console.log('%c[REPLACE] Starting search and replace...', 'color: #2196f3; font-weight: bold;');
        console.log('[REPLACE] Original search text:', searchText.substring(0, 100) + '...');
        console.log('[REPLACE] Original replace text:', replaceText.substring(0, 100) + '...');

        // Check if this is a multi-line replacement (contains \r\n or \n)
        const hasMultipleLines = searchText.includes('\r\n') || searchText.includes('\n');

        if (hasMultipleLines) {
            console.log('%c[REPLACE] Multi-line replacement detected', 'color: #ff9800; font-weight: bold;');

            // Split into individual lines
            const searchLines = searchText.split(/\r\n|\n/).filter(function(line) { return line.trim(); });
            const replaceLines = replaceText.split(/\r\n|\n/).filter(function(line) { return line.trim(); });

            console.log('[REPLACE] Search lines:', searchLines);
            console.log('[REPLACE] Replace lines:', replaceLines);

            // If same number of lines, do individual replacements
            if (searchLines.length === replaceLines.length) {
                console.log('[REPLACE] Performing individual line replacements...');
                searchLines.forEach(function(searchLine, index) {
                    const replaceLine = replaceLines[index];
                    console.log('[REPLACE] Line ' + (index + 1) + ': "' + searchLine.trim() + '" → "' + replaceLine.trim() + '"');

                    window.Asc.plugin.executeMethod("SearchAndReplace", [{
                        searchString: searchLine.trim(),
                        replaceString: replaceLine.trim(),
                        matchCase: false
                    }]);
                });
            } else {
                // Different number of lines - try as block with normalized whitespace
                console.log('[REPLACE] Different line counts, trying block replacement...');
                const cleanSearchText = searchText.replace(/\r\n/g, '\n').replace(/\s+/g, ' ').trim();
                const cleanReplaceText = replaceText.replace(/\r\n/g, '\n').replace(/\s+/g, ' ').trim();

                window.Asc.plugin.executeMethod("SearchAndReplace", [{
                    searchString: cleanSearchText,
                    replaceString: cleanReplaceText,
                    matchCase: false
                }]);
            }
        } else {
            // Single line replacement
            console.log('[REPLACE] Single line replacement');
            const cleanSearchText = searchText.replace(/\s+/g, ' ').trim();
            const cleanReplaceText = replaceText.replace(/\s+/g, ' ').trim();

            console.log('[REPLACE] Cleaned search:', cleanSearchText);
            console.log('[REPLACE] Cleaned replace:', cleanReplaceText);

            window.Asc.plugin.executeMethod("SearchAndReplace", [{
                searchString: cleanSearchText,
                replaceString: cleanReplaceText,
                matchCase: false
            }]);
        }

        console.log('%c[REPLACE] Search and replace commands sent', 'color: #4caf50; font-weight: bold;');
    }

    function tryFuzzyReplace(targetText, newText) {
        window.Asc.plugin.executeMethod("GetAllParagraphs", [], function(paragraphs) {
            if (!paragraphs) {
                console.log('No paragraphs available for fuzzy matching');
                return;
            }

            let bestMatch = null;
            let bestSimilarity = 0;

            paragraphs.forEach(function(para, index) {
                if (!para || !para.Text) return;

                const similarity = calculateSimilarity(targetText, para.Text.trim());

                if (similarity >= 90 && similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                    bestMatch = {
                        text: para.Text.trim(),
                        index: index,
                        similarity: similarity
                    };
                }
            });

            if (bestMatch) {
                console.log('Found fuzzy match with ' + bestSimilarity.toFixed(1) + '% similarity:', bestMatch.text.substring(0, 50) + '...');

                // Replace the fuzzy matched text
                window.Asc.plugin.executeMethod("SearchAndReplace", [{
                    searchString: bestMatch.text,
                    replaceString: newText,
                    matchCase: false
                }]);
            } else {
                console.log('No match found with >90% similarity for:', targetText.substring(0, 50) + '...');
            }
        });
    }

    function showMessage(type, content) {
        const messageDiv = document.createElement('div');
        messageDiv.className = 'message ' + type + '-message';

        const messageContent = document.createElement('div');
        messageContent.className = 'message-content';
        messageContent.textContent = content;

        const timestamp = document.createElement('div');
        timestamp.className = 'message-time';
        timestamp.textContent = new Date().toLocaleTimeString([], {
            hour: '2-digit',
            minute: '2-digit'
        });

        messageDiv.appendChild(messageContent);
        messageDiv.appendChild(timestamp);

        chatContainer.appendChild(messageDiv);

        // Scroll to bottom
        chatContainer.scrollTop = chatContainer.scrollHeight;

        // Auto-remove system messages after 5 seconds
        if (type === 'system') {
            setTimeout(function() {
                if (messageDiv.parentNode) {
                    messageDiv.style.transition = 'opacity 0.5s ease-out';
                    messageDiv.style.opacity = '0';
                    setTimeout(function() {
                        if (messageDiv.parentNode) {
                            messageDiv.parentNode.removeChild(messageDiv);
                        }
                    }, 500);
                }
            }, 5000);
        }
    }

    function updateChatCount() {
        chatCount.textContent = 'Messages: ' + messageCount + '/' + MAX_MESSAGES;

        if (messageCount >= MAX_MESSAGES) {
            chatCount.style.color = '#d32f2f';
            sendBtn.disabled = true;
            sendBtn.textContent = 'Limit Reached';
        } else {
            chatCount.style.color = '';
        }
    }

    function updateImageCount() {
        chatCount.textContent = 'Images: ' + imageGenerationCount + '/' + MAX_IMAGE_GENERATIONS;

        if (imageGenerationCount >= MAX_IMAGE_GENERATIONS) {
            chatCount.style.color = '#d32f2f';
        } else {
            chatCount.style.color = '';
        }
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function formatNumber(num) {
        return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    }

    window.Asc.plugin.onExternalMouseUp = function () {
        var evt = document.createEvent("MouseEvents");
        evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0, false, false, false, false, 0, null);
        document.dispatchEvent(evt);
    };

    window.Asc.plugin.onContextMenuShow = function(options) {
        // Only show menu item if text is selected
        if (options && options.type === "Selection") {
            return;
        }
    };

    window.Asc.plugin.onContextMenuClick = function(id) {
        if (id === "attachToRFPChatbot") {
            // Get the selected text and attach it as context
            window.Asc.plugin.executeMethod("GetSelectedText", [true], function(text) {
                if (!text || text.trim() === '') {
                    showMessage('system', 'No text selected.');
                    return;
                }

                const contextKey = Object.keys(currentContext).length + 1;

                if (contextKey > 2) {
                    showMessage('system', 'Maximum 2 contexts can be attached at once.');
                    return;
                }

                // Check character limit
                const contentLength = text.length;
                const remainingChars = MAX_CONTEXT_CHARS - totalContextChars;

                if (contentLength > MAX_CONTEXT_CHARS) {
                    showMessage('system', 'Context too long! Selected text has ' + formatNumber(contentLength) + ' characters. Maximum allowed is ' + formatNumber(MAX_CONTEXT_CHARS) + ' characters.');
                    return;
                }

                if (contentLength > remainingChars) {
                    showMessage('system', 'Context too long! You have ' + formatNumber(totalContextChars) + ' characters already. Selected text (' + formatNumber(contentLength) + ' chars) exceeds the remaining limit of ' + formatNumber(remainingChars) + ' characters.');
                    return;
                }

                currentContext[contextKey] = {
                    content: text,
                    isHeader: false,
                    charCount: contentLength
                };
                selectedText = text;
                totalContextChars += contentLength;

                // Highlight the selected text in the document
                highlightSelectedText();

                // Update context display
                updateContextDisplay();

                // Focus on the message input
                setTimeout(function() {
                    messageInput.focus();
                }, 100);
            });
        }
    };

    window.Asc.plugin.button = function(id) {
        console.log('Plugin button clicked:', id);

        if (id === -1) {
            this.executeCommand("close", "");
        }
    };

    window.Asc.plugin.event_onToolbarMenuClick = function(id) {
        console.log('Toolbar menu clicked:', id);

        if (id === "openChatbot") {
            // Show the chatbot panel
            var pluginContainer = document.querySelector('.plugin-container');
            if (pluginContainer) {
                pluginContainer.style.display = 'flex';
            }
            var imageGenContainer = document.getElementById('imageGeneratorContainer');
            if (imageGenContainer) {
                imageGenContainer.style.display = 'none';
            }
            if (messageInput) {
                messageInput.focus();
            }
        } else if (id === "openImageGenerator") {
            // Open image generator interface
            openImageGenerator();
        }
    };

    function openImageGenerator() {
        // Hide chatbot interface
        var mainContainer = document.querySelector('.plugin-container');
        if (mainContainer) {
            mainContainer.style.display = 'none';
        }

        // Create or show image generator interface
        let imageGenContainer = document.getElementById('imageGeneratorContainer');

        if (!imageGenContainer) {
            imageGenContainer = document.createElement('div');
            imageGenContainer.id = 'imageGeneratorContainer';
            imageGenContainer.className = 'plugin-container';
            imageGenContainer.innerHTML = `
                <div class="header">
                    <h2>Image Generator</h2>
                    <button id="backToChatbot" style="background: #4CAF50; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;">← Back to Chatbot</button>
                </div>

                <div class="chat-container" id="imageGenMessages" style="flex: 1; overflow-y: auto; padding: 12px;">
                    <div class="welcome-message">
                        <p>Describe the image you want to generate for your document.</p>
                    </div>
                </div>

                <div class="input-container">
                    <textarea
                        id="imagePromptInput"
                        placeholder="Describe the image you want to generate..."
                        rows="3"
                    ></textarea>
                    <div class="input-actions">
                        <button class="send-btn" id="generateImageBtn">Generate</button>
                    </div>
                </div>

                <div style="padding: 8px 12px; background: #f9f9f9; border-top: 1px solid #e0e0e0; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0;">
                    <span class="chat-count" id="imageGenCount" style="font-size: 11px; color: #666; font-weight: 500;">Generations: 0/2</span>
                    <button id="resetImageGeneration" style="background: #ff9800; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 11px; font-weight: 600;">🔄 Start New</button>
                </div>
            `;
            document.body.appendChild(imageGenContainer);

            // Add event listeners
            document.getElementById('backToChatbot').addEventListener('click', function() {
                imageGenContainer.style.display = 'none';
                if (mainContainer) {
                    mainContainer.style.display = 'flex';
                }
            });

            document.getElementById('resetImageGeneration').addEventListener('click', function() {
                resetImageGeneration();
            });

            document.getElementById('generateImageBtn').addEventListener('click', function() {
                handleImageGeneration();
            });

            document.getElementById('imagePromptInput').addEventListener('keydown', function(e) {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    handleImageGeneration();
                }
            });
        } else {
            imageGenContainer.style.display = 'flex';
        }
    }

    function resetImageGeneration() {
        currentImageCallId = null;
        currentConversationId = null;
        imageGenerationCount = 0;

        // Clear messages
        const messagesContainer = document.getElementById('imageGenMessages');
        if (messagesContainer) {
            messagesContainer.innerHTML = `
                <div class="welcome-message">
                    <p>🎨 Welcome to Image Generator!</p>
                    <p>Describe the image you want to generate for your RFP/Grant document.</p>
                    <p style="margin-top: 8px; font-size: 12px; color: #666;">You can generate up to 2 images per session.</p>
                </div>
            `;
        }

        // Reset counter
        updateImageGenCount();

        // Re-enable button
        const generateBtn = document.getElementById('generateImageBtn');
        if (generateBtn) {
            generateBtn.disabled = false;
            generateBtn.textContent = 'Generate';
        }

        showImageGenMessage('system', '✓ Image generation reset. You can start a new generation session.');
    }


})(window, undefined);
