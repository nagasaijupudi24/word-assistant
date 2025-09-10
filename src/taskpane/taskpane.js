import { mark } from "regenerator-runtime";
import { model4one, model4 } from "../../webconfig.js";

const CHUNK_SIZE = 1200;
const CHUNK_OVERLAP = 200;
const MAX_CHUNKS = 3;
const MAX_HISTORY = 10;

// Global State
let documentChunks = [];
let chunkOffsets = [];
let conversationHistory = [];
let isStreaming = false;
let isProcessing = false;
let abortController = null;
let currentDocumentContent = "";
let documentContentAvailable = false;
let isVisible = false;
let newChatButton = document.getElementById("newChatButton");
let chatHistory = document.getElementById("chatHistory");

let toneValue = document.getElementById("selectedToneLabel");
// console.log(toneValue.textContent,'Selected Tone Label')

let selectedModel = document.getElementById("selectedModelLabel");
console.log(selectedModel.textContent, "Selected Model Label");


const promptsContainer = document.getElementById("prompt-suggestions-container");

let modelCredentials = model4one;


// UI State
let lastUserMessage = null;
let lastAIResponse = null;

// Initialize when Office is ready
if (window.Office && Office.onReady) {
  Office.onReady(() => {
    initializeApp();

  });
} else {
  document.addEventListener("DOMContentLoaded", initializeApp);
}

function initializeApp() {
  setupChatInterface();
  setupDocumentContentMonitoring();
  loadConversation();
  setupUIEventListeners();
}

function setupUIEventListeners() {
  // Prompt suggestions toggle
  const btnPrompt = document.getElementById("suggestions");
  


btnPrompt?.addEventListener("click", () => {
  isVisible = !isVisible; // Toggle the state
  promptsContainer?.classList.toggle("hide", !isVisible);
  console.log("Prompt suggestions button clicked");
  console.log(isVisible, 'isVisible in btnPrompt');
});

  // Suggestion pills
  document.querySelectorAll(".suggestion-pill").forEach((btn) => {
    btn.addEventListener("click", function () {
      const input = document.getElementById("userPrompt");
      if (input) {
        input.value = this.getAttribute("data-suggestion");
        input.focus();
        promptsContainer?.classList.add("hide");
        ""
          console.log(isVisible, 'isVisible in suggestion pill');
        isVisible = false;
      }
    });
  });

  // Settings and tone dropdown
  const settingsBtn = document.getElementById("settingsButton");
  const settingsTooltip = document.getElementById("settingsTooltip");
  const toneDropdownBtn = document.getElementById("toneDropdownBtn");
  const toneOptions = document.getElementById("toneOptions");
  const selectedToneLabel = document.getElementById("selectedToneLabel");

  let isSettingsButtonVisible = false;

  settingsBtn.addEventListener("click", (e) => {
    console.log("Settings button clicked");
    e.stopPropagation();
    if (isSettingsButtonVisible) {
      settingsTooltip.classList.add("hide");
      isSettingsButtonVisible = false;
    } else {
      settingsTooltip.classList.remove("hide");
      isSettingsButtonVisible = true;
    }
  });
  toneDropdownBtn?.addEventListener("click", (e) => {
    e.stopPropagation();
    toneOptions?.classList.remove("hide");
  });

  toneOptions?.querySelectorAll(".tone-option-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedToneLabel.textContent = btn.getAttribute("data-value");
      toneOptions?.classList.add("hide");
       console.log(toneValue.textContent,'Selected Tone Label after click')
      setTimeout(() => settingsTooltip?.classList.add("hide"), 200);
    });
  });

  const modelDropdownBtn = document.getElementById("modelDropdownBtn");
  const modelOptions = document.getElementById("modelOptions");
  const selectedModelLabel = document.getElementById("selectedModelLabel");

  modelDropdownBtn?.addEventListener("click", (e) => {
    e.stopPropagation();
    modelOptions?.classList.remove("hide");
  });

  modelOptions?.querySelectorAll(".modelSelector").forEach((modelBtn) => {
    modelBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedModelLabel.textContent = modelBtn.getAttribute("data-value");
      modelOptions?.classList.add("hide");
      switch (selectedModelLabel.textContent) {
        case "GPT 4.1":
          modelCredentials = model4one;
          break;
        case "GPT 4.1 mini":
          modelCredentials = model4;
      }
      isStreaming = false; // <-- Add this line to fix the issue
      updateSendButton(false); // Optionally reset button UI
      console.log(modelCredentials, "Model Credentials after selection");
      console.log(selectedModel.textContent, "Selected Model Label after click");
      setTimeout(() => document.getElementById("settingsTooltip")?.classList.add("hide"), 1000);
    });
  });

  // About modal
  const aboutBtn = document.getElementById("about-btn");
  const aboutModal = document.getElementById("about-modal");
  const aboutModalClose = document.getElementById("about-modal-close");

  aboutBtn?.addEventListener("click", () => {
    aboutModal?.classList.add("show");
    settingsTooltip?.classList.add("hide");
  });

  aboutModalClose?.addEventListener("click", () => {
    aboutModal?.classList.remove("show");
  });

  aboutModal?.addEventListener("click", (e) => {
    if (e.target === aboutModal) {
      aboutModal.classList.remove("show");
    }
  });

  // Close all dropdowns when clicking outside
  document.addEventListener("click", () => {
    settingsTooltip?.classList.add("hide");
    toneOptions?.classList.add("hide");
  });
}

function setupChatInterface() {
  const sendButton = document.getElementById("send-button");
  const userPrompt = document.getElementById("userPrompt");
  const newChatButton = document.getElementById("newChatButton");

  userPrompt?.addEventListener("input", function () {
    this.style.height = "auto";
    this.style.height = this.scrollHeight + "px";
  });

  sendButton?.addEventListener("click", function () {
    isStreaming ? stopStreaming() : sendMessage();
  });

  userPrompt?.addEventListener("keydown", function (event) {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      sendMessage();
    }
  });

  userPrompt?.addEventListener("focus", () => {
  promptsContainer?.classList.add("hide");
  isVisible = false;
});

  newChatButton?.addEventListener("click", startNewChat);

  document.querySelectorAll(".example-card").forEach((card) => {
    card.addEventListener("click", function () {
      const prompt = this.getAttribute("data-prompt");
      const userPrompt = document.getElementById("userPrompt");
      if (userPrompt) {
        userPrompt.value = prompt;
        userPrompt.dispatchEvent(new Event("input"));
        setTimeout(sendMessage, 100);
      }
    });
  });
}

// Document Processing
function chunkDocument(content) {
  documentChunks = [];
  chunkOffsets = [];
  if (!content) return [];

  let pos = 0;
  while (pos < content.length) {
    let end = Math.min(pos + CHUNK_SIZE, content.length);

    // Try to end at sentence boundary
    const sentenceBoundary = Math.max(
      content.lastIndexOf(".", end),
      content.lastIndexOf("?", end),
      content.lastIndexOf("!", end),
      content.lastIndexOf("\n\n", end)
    );

    if (sentenceBoundary > pos + CHUNK_SIZE / 2) {
      end = sentenceBoundary + 1;
    }

    const chunk = content.substring(pos, end).trim();
    if (chunk) {
      documentChunks.push(chunk);
      chunkOffsets.push({ start: pos, end });
    }

    pos = Math.max(end - CHUNK_OVERLAP, pos + 1);
    if (pos >= content.length) break;
  }

  // document.getElementById("chunkCount").textContent = documentChunks.length;
  return documentChunks;
}

function getRelevantChunks(query, maxChunks = MAX_CHUNKS) {
  console.log(query)
  debugger
  if (!documentChunks.length) return [];

  const keywords = query
    .toLowerCase()
    .split(/\W+/)
    .filter((x) => x.length > 2);
  const scores = documentChunks.map((chunk, idx) => {
    const text = chunk.toLowerCase();
    let score = 0;
    keywords.forEach((word) => {
      score += (text.match(new RegExp(`\\b${word}\\b`, "g")) || []).length * 2;
      score += (text.match(new RegExp(word, "g")) || []).length;
    });
    return { idx, score };
  });

  scores.sort((a, b) => b.score - a.score);
  const selected = new Set(scores.slice(0, maxChunks).map((s) => s.idx));

  // Always include first and last chunks
  selected.add(0);
  selected.add(documentChunks.length - 1);

  return Array.from(selected)
    .sort((a, b) => a - b)
    .map((i) => documentChunks[i]);
}

// Document Content Handling
function setupDocumentContentMonitoring() {
  getFullDocumentContent().then(handleDocumentContent);

  // Listen for document content changes and chunk updated content
  if (window.Word && Office.context?.document) {
    // Office.js: Listen for document body changes
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async function () {
        // Get updated content and chunk
        const updatedContent = await getFullDocumentContent();
        chunkDocument(updatedContent);
      }
    );
  } else {
    // Fallback: Simulate with a timer for demo/browser
    setInterval(async () => {
      const updatedContent = await getFullDocumentContent();
      if (updatedContent !== currentDocumentContent) {
        chunkDocument(updatedContent);
        currentDocumentContent = updatedContent;
      }
    }, 2000); // Check every 2 seconds
  }
}

async function getFullDocumentContent() {
  return new Promise((resolve) => {
    if (window.Word) {
      Word.run((context) => {
        const body = context.document.body;
        body.load("text");
        return context
          .sync()
          .then(() => resolve(body.text))
          .catch(() => resolve(currentDocumentContent || ""));
      }).catch(() => resolve(currentDocumentContent || ""));
    } else {
      resolve(
        currentDocumentContent ||
          "This is simulated document content. In a real implementation, this would be retrieved from the Word document."
      );
    }
  });
}

function handleDocumentContent(content) {
  currentDocumentContent = content || "";
  documentContentAvailable = !!content?.trim();

  // if (!documentContentAvailable) {
  //    newChatButton.disabled = true;
  //   showNullDocumentMessage();
  //   return;
  // }

  chunkDocument(content);
  enableInput();

  const docPreview = document.getElementById("documentContent");
  if (docPreview) {
    docPreview.textContent = content.length > 1000 ? content.substring(0, 1000) + "..." : content;
    docPreview.classList.remove("empty");
  }
}

function showNullDocumentMessage() {
  const chatHistory = document.getElementById("chatHistory");
  if (chatHistory) {
    chatHistory.innerHTML = `   
      <div class="welcome-section">   
        <div class="welcome-icon">   
          <i class="fas fa-exclamation-circle"></i>   
        </div>   
        <h2 class="welcome-title">No Document Content</h2>   
        <p class="welcome-subtitle">Document content is <b>empty</b>. Please add text to use chat features.</p>   
      </div>`;
  }
  disableInput();
}

function disableInput() {
  const input = document.getElementById("userPrompt");
  const button = document.getElementById("send-button");
  if (input) input.disabled = true;
  if (button) button.disabled = true;
}

function enableInput() {
  document.getElementById("chatHistory").classList.remove("hide");
  const input = document.getElementById("userPrompt");
  const button = document.getElementById("send-button");
  newChatButton.disabled = false;
  if (input) input.disabled = false;
  if (button) button.disabled = false;
}

// Conversation Management
function saveConversation() {
  document.getElementById("chatHistory").classList.remove("welcome-contianer");
  try {
    const trimmed = conversationHistory.slice(-MAX_HISTORY);
    localStorage.setItem("documentChatHistory", JSON.stringify(trimmed));

    // Add null check for historyCount element
    const historyCountElem = document.getElementById("historyCount");
    if (historyCountElem) {
      historyCountElem.textContent = trimmed.length;
    }
  } catch (e) {
    console.error("Error saving conversation:", e);
  }
}

function loadConversation() {
  try {
    const saved = localStorage.getItem("documentChatHistory");
    let loaded = [];
    if (saved) {
      // Filter out any system messages and duplicates
      loaded = JSON.parse(saved).slice(-MAX_HISTORY);
      loaded = loaded.filter((msg, idx, arr) => {
        // Remove system messages
        if (msg.role === "system") return false;
        // Remove consecutive duplicates
        if (idx > 0 && msg.role === arr[idx - 1].role && msg.content === arr[idx - 1].content)
          return false;
        return true;
      });
    }
    conversationHistory = loaded;
    if (conversationHistory.length === 0) {
      showWelcomeSection();
    } else {
      renderConversation();
    }
  } catch (e) {
    console.error("Error loading conversation:", e);
  }

  // Show welcome section if chat history is empty
  function showWelcomeSection() {
    if (chatHistory) {
      chatHistory.classList.add("welcome-contianer");
      chatHistory.style.justifyContent = "center";
      chatHistory.innerHTML = `   
      <div id="welcome-section">
        <div class="welcome-header">
          <img src="../../assets/logo-filled.png" alt="Copilot">
          <h2>How can I help with this Word?</h2>
        </div>
      </div>`;
    }
  }
}

function renderConversation() {
  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return;
  // Remove welcome section if present
  const welcomeSection = document.getElementById("welcome-section");
  if (welcomeSection) welcomeSection.remove();
  chatHistory.classList.remove("welcome-contianer");
  chatHistory.innerHTML = "";
  chatHistory.style.justifyContent = "flex-start";
  console.log(conversationHistory, "Conversation History");
  conversationHistory.forEach((msg) => {
    if (msg.role === "user" || msg.role === "assistant") {
      addMessageToUI(msg.role, msg.content);
    }
  });
}

function startNewChat() {
  promptsContainer?.classList.add("hide");
  isVisible = false;
  
  conversationHistory = [];
  lastUserMessage = null;
  lastAIResponse = null;
  saveConversation();
  const chatHistory = document.getElementById("chatHistory");
  chatHistory.style.justifyContent = "center";
  if (chatHistory) {
    chatHistory.innerHTML = `   
      <div id="welcome-section">
        <div class="welcome-header">
          <img src="../../assets/logo-filled.png" alt="Copilot">
          <h2>How can I help with this Word?</h2>
        </div>
      </div>`;
  }
}

// Message Handling
async function sendMessage() {
  if (isProcessing) return;
  isProcessing = true;

  const userPrompt = document.getElementById("userPrompt");
  const message = userPrompt?.value.trim() || "";

  if (!message || message === lastUserMessage) {
    isProcessing = false;
    return;
  }
  lastUserMessage = message;

  // Clear input and prepare UI
  if (userPrompt) {
    userPrompt.value = "";
    userPrompt.style.height = "auto";
  }

  // Remove welcome section and update container style
  let welcomeSection = document.getElementById("welcome-section");
  if (welcomeSection) {
    welcomeSection.remove();
    document.getElementById("chatHistory")?.classList.remove("welcome-contianer");
  }

  // Add user message immediately
  addMessageToUI("user", message);
  conversationHistory.push({ role: "user", content: message });

  // Show typing indicator
  addTypingIndicator();
  updateSendButton(true);

  try {
    // Get document context
    currentDocumentContent = await getFullDocumentContent();
    documentContentAvailable = !!currentDocumentContent?.trim();

    // if (!documentContentAvailable) {
    //   showNullDocumentMessage();
    //   return;
    // }

    // Process message
    const relevantChunks = getRelevantChunks(message, MAX_CHUNKS);
    console.log(relevantChunks, "Relevant Chunks");
    const messages = createMessagePayload(message, relevantChunks);
    console.log(messages, "Messages sent to AI");

    // Get AI response
    await getAIResponse(messages);
    saveConversation();
  } catch (error) {
    console.error("Error in sendMessage:", error);
    const errorMessage = `Sorry, I encountered an error: ${error.message || "Please try again"}`;
    addMessageToUI("assistant", errorMessage);
    conversationHistory.push({ role: "assistant", content: errorMessage });
    saveConversation();
  } finally {
    removeTypingIndicator();
    updateSendButton(false);
    isProcessing = false;
  }
}

function createMessagePayload(userMessage, relevantChunks) {
  const contextContent = relevantChunks.join("\n\n---\n\n");
  const systemMessage = {
    role: "system",
    content: `You are an AI assistant for Word documents. Use ONLY the provided context to answer questions or generate content. Always answer, even if the relevant info is at the very end or start. Respond in Markdown.   
        
    You must always respond with valid Format in the following format:
 

  the primary response text here
  "SEP"
  "follow_up_questions": [
    "<question 1>",
    "<question 2>",
    "<question 3>"
  ]

      Relevant document context:   
      ${contextContent}.
      
       Always respond in ${toneValue ? toneValue.textContent : ""}.`,
  };

  const mappedHistory = conversationHistory
    .slice(-MAX_HISTORY)
    .map(({ role, content }) => ({ role, content }));

  return [systemMessage, ...mappedHistory];
}




function updateSuggestionPills(questions) {
  const container = document.getElementById("prompt-suggestions");
  if (!container) return;

  // Clear old pills
  container.innerHTML = "";

  // Add new pills
  questions.forEach((q) => {
    const btn = document.createElement("button");
    btn.className = "suggestion-pill";
    btn.setAttribute("data-suggestion", q);
    btn.textContent = q;

    btn.addEventListener("click", () => {
      const input = document.getElementById("userPrompt");
      if (input) {
        input.value = q;
        input.focus();
        promptsContainer.classList.add("hide");
        isVisible = false; // <-- Ensure state is updated!
      }
    });

    container.appendChild(btn);
  });

  // Make container visible if it was hidden
  const outer = document.getElementById("prompt-suggestions-container");
   outer.classList.remove("hide");
    isVisible = true; // <-- Ensure state is updated!
    console.log(isVisible, 'isVisible in updateSuggestionPills');
 
}


async function getAIResponse(messages) {
  const fetchAbortController = new AbortController();
  abortController = fetchAbortController;
  isStreaming = true;

  let accumulatedResponse = "";
  let mainResponse = "";
  let follow_up_questions;
  let messageElement = null;



  try {
    const response = await fetch(
      `${modelCredentials.AZURE_ENDPOINT}openai/deployments/${modelCredentials.AZURE_DEPLOYMENT}/chat/completions?api-version=${modelCredentials.AZURE_API_VERSION}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": modelCredentials.AZURE_API_KEY,
        },
        body: JSON.stringify({
          messages,
          max_tokens: 800,
          temperature: 0.7,
          top_p: 0.95,
          stream: true,
        }),
        signal: fetchAbortController.signal,
      }
    );

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop() || "";

      for (const line of lines) {
        if (!line.startsWith("data: ")) continue;

        const dataStr = line.substring(6).trim();
        if (dataStr === "[DONE]") {
          // console.log("Stream complete");
          // console.log("Final mainResponse response:",  accumulatedResponse.split("SEP")[0].trim());
          let trimmedResponse = accumulatedResponse.split("SEP")[0].trim();
          finalizeResponse(trimmedResponse);
          return;
        }

        try {
          const data = JSON.parse(dataStr);
          const content = data.choices[0]?.delta?.content || "";

          if (content) {
            accumulatedResponse += content;

            // console.log("Delta received:", content);
            // console.log("Type of delta:", typeof content);

            if (content.includes("SEP")) {
              console.log("Substring found!");
            }

            let [mainResponse, followupPart] = accumulatedResponse.split("SEP");
            // console.log("Main response so far:", mainResponse);
            mainResponse = mainResponse.trim();
            // console.log("Follow-up part so far:", followupPart);
            // console.log("Type of followupPart:", typeof followupPart);

            if (followupPart && typeof followupPart === "string") {
              let cleaned = followupPart.trim();
              if (cleaned.startsWith('"follow_up_questions"')) {
                cleaned = `{ ${cleaned} }`;
              }
              try {
                const parsed = JSON.parse(cleaned);
                follow_up_questions = parsed.follow_up_questions || [];

                
                // console.log("Parsed follow-up questions:", follow_up_questions);
                 // ✅ Replace suggestion pills dynamically
            updateSuggestionPills(follow_up_questions);
              } catch (e) {
                console.error("Error parsing follow-up questions JSON:", e);
              }
            }
            // console.log("Final follow-up questions:", follow_up_questions);

            // console.log("Accumulated response so far:", accumulatedResponse);

            messageElement = updateOrCreateMessage(messageElement, mainResponse);

            
          }
        } catch (e) {
          console.error("Error parsing JSON:", e);
        }
      }
    }

  
  } catch (error) {
    handleAIError(error, mainResponse);
  } finally {
    // ""
    cleanupResponse(messageElement, mainResponse);
  }
}

function updateOrCreateMessage(messageElement, content) {
  if (!messageElement) {
    removeTypingIndicator();
    return addMessageToUI("assistant", content, true);
  }
  document.getElementById("chatHistory").style.justifyContent = "flex-start";

  const contentElement = messageElement.querySelector(".message-content");
  if (contentElement) {
    contentElement.innerHTML = marked.parse(content);
    messageElement.scrollIntoView({ behavior: "smooth", block: "end" });
  }
  return messageElement;
}

function finalizeResponse(content) {
  conversationHistory.push({ role: "assistant", content });
  saveConversation();
  isStreaming = false;
}

function handleAIError(error, accumulatedResponse) {
  if (error.name === "AbortError") {
    console.log("Stream aborted by user");
    if (accumulatedResponse) {
      conversationHistory.push({ role: "assistant", content: accumulatedResponse });
    }
  } else {
    console.error("Streaming error:", error);
    const errorMessage = `Error: ${error.message || "Failed to get response"}`;
    addMessageToUI("assistant", errorMessage);
    conversationHistory.push({ role: "assistant", content: errorMessage });
  }
}

function cleanupResponse(messageElement, accumulatedResponse) {
  isStreaming = false;
  removeTypingIndicator();
  updateSendButton(false);

  if (accumulatedResponse && messageElement) {
    const chatHistory = document.getElementById("chatHistory");
    if (chatHistory) {
      messageElement.remove();
      addMessageToUI("assistant", accumulatedResponse,isStreaming);
    }
  }
}

function stopStreaming() {
  if (isStreaming && abortController) {
    abortController.abort();
    isStreaming = false;
    updateSendButton(false);
    removeTypingIndicator();
  }
}

function addMessageToUI(role, content, isStreaming = false) {
  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return null;
  // ""

  if (role === "assistant") {
    removeTypingIndicator();
  }

  const messageElement = document.createElement("div");
  messageElement.className = `message ${role === "assistant" ? "ai-message" : "user-message"}`;

  const header = document.createElement("div");
  header.className = "message-header";
  header.innerHTML = `
    <span>
      <i class="fas ${role === "user" ? "fa-user" : "fa-robot"}"></i>
      ${role === "user" ? "You" : "Document Assistant"}
    </span>
    <span>${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}</span>
  `;

  const contentElement = document.createElement("div");
  contentElement.className = "message-content";
  contentElement.innerHTML = marked.parse(content);

  messageElement.appendChild(header);
  messageElement.appendChild(contentElement);
if (role === "assistant") {
  const actions = document.createElement("div");
  actions.className = "message-actions";

  // Insert button
  const insertButton = document.createElement("button");
  insertButton.className = "action-button insert-button";
  insertButton.innerHTML = '<i class="fas fa-plus"></i> Insert into Document';
  insertButton.onclick = () => {
    const currentContent = contentElement.innerHTML; // get latest rendered content
    insertContent(currentContent);
  };
  actions.appendChild(insertButton);

  // Copy button
  const copyButton = document.createElement("button");
  copyButton.className = "action-button copy-button";
  copyButton.innerHTML = '<i class="far fa-copy"></i> Copy';
  copyButton.onclick = (e) => {
    const currentContent = contentElement.innerHTML; // plain text for clipboard
    copyToClipboard(currentContent, e);
  };
  actions.appendChild(copyButton);

  // 👍 Like button
  const likeButton = document.createElement("button");
  likeButton.className = "action-button like-button";
  likeButton.innerHTML = '<i class="far fa-thumbs-up"></i>'; // outline by default
  likeButton.onclick = () => {
    const icon = likeButton.querySelector("i");
    if (icon.classList.contains("far")) {
      icon.classList.replace("far", "fas"); // fill
      console.log('Like button clicked');
    } else {
      icon.classList.replace("fas", "far"); // unfill
      console.log('Like button unclicked');
    }
  };
  actions.appendChild(likeButton);

  // 👎 Dislike button
  const dislikeButton = document.createElement("button");
  dislikeButton.className = "action-button dislike-button";
  dislikeButton.innerHTML = '<i class="far fa-thumbs-down"></i>'; // outline by default
  dislikeButton.onclick = () => {
    const icon = dislikeButton.querySelector("i");
    if (icon.classList.contains("far")) {
      icon.classList.replace("far", "fas"); // fill
      console.log('Dislike button clicked');
    } else {
      icon.classList.replace("fas", "far"); // unfill
      console.log('Dislike button unclicked');
    }
  };
  actions.appendChild(dislikeButton);

  messageElement.appendChild(actions);
}




  chatHistory.appendChild(messageElement);
  messageElement.scrollIntoView({ behavior: "smooth", block: "end" });
  return messageElement;
}

// function createActionButton(className, html, onClick) {
//   const button = document.createElement("button");
//   button.className = `action-button ${className}`;
//   button.innerHTML = html;
//   button.onclick = onClick;
//   return button;
// }

function addTypingIndicator() {
  removeTypingIndicator();
  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return;

  const typingElement = document.createElement("div");
  typingElement.className = "message ai-message";
  typingElement.id = "typingIndicator";

  const header = document.createElement("div");
  header.className = "message-header";
  header.innerHTML =
    '<span><i class="fas fa-robot"></i> Document Assistant</span><span>Typing...</span>';

  const indicator = document.createElement("div");
  indicator.className = "typing-indicator";
  indicator.innerHTML = `   
    <div class="typing-dot"></div>   
    <div class="typing-dot"></div>   
    <div class="typing-dot"></div>   
    <span class="typing-text">Thinking...</span>`;

  typingElement.appendChild(header);
  typingElement.appendChild(indicator);
  chatHistory.appendChild(typingElement);
  typingElement.scrollIntoView({ behavior: "smooth", block: "end" });
}

function removeTypingIndicator() {
  document.getElementById("typingIndicator")?.remove();
}

function updateSendButton(isStreaming) {
  const sendButton = document.getElementById("send-button");
  if (!sendButton) return;

  sendButton.classList.toggle("stop-button", isStreaming);
  sendButton.innerHTML = isStreaming
    ? '<i class="fas fa-stop"></i>'
    : '<i class="fas fa-paper-plane"></i>';
}

// Content Actions
function isDraftContent(content) {
  if (!content) return false;
  const draftKeywords = [
    "draft",
    "summary",
    "outline",
    "content",
    "section",
    "conclusion",
    "introduction",
    "generated",
    "paragraph",
    "bullet points",
    "points",
  ];
  const lower = content.toLowerCase();
  return draftKeywords.some((k) => lower.includes(k)) || content.length > 100;
}

function insertContent(content) {
  // Always use HTML formatting for both insert and copy actions
  let htmlContent = content;
  ""
  if (!/<[a-z][\s\S]*>/i.test(content)) {
    htmlContent = markdownToHtml(content);
  }
  if (window.Word && Office.context?.document) {
    Word.run(async (context) => {
      const range = context.document.getSelection();
      try {
        range.insertHtml(htmlContent, Word.InsertLocation.replace);
      } catch (e) {
        range.insertText(content, Word.InsertLocation.replace);
      }
      await context.sync();
    }).catch((error) => {
      console.error("Error inserting content:", error);
      alert("Failed to insert content. Please try again.");
    });
  } else {
    // For browser/demo, just log the HTML
    console.log("Word API not available. This feature works in Word as an Office Add-in.");
    console.log("HTML to insert:", htmlContent);
  }
}

// Simple Markdown to HTML converter for basic formatting
function markdownToHtml(md) {
  let html = md;
  // Headings
  html = html.replace(/^###### (.*$)/gim, "<h6>$1</h6>");
  html = html.replace(/^##### (.*$)/gim, "<h5>$1</h5>");
  html = html.replace(/^#### (.*$)/gim, "<h4>$1</h4>");
  html = html.replace(/^### (.*$)/gim, "<h3>$1</h3>");
  html = html.replace(/^## (.*$)/gim, "<h2>$1</h2>");
  html = html.replace(/^# (.*$)/gim, "<h1>$1</h1>");
  // Bold
  html = html.replace(/\*\*(.*?)\*\*/gim, "<b>$1</b>");
  // Italic
  html = html.replace(/\*(.*?)\*/gim, "<i>$1</i>");
  // Lists
  html = html.replace(/^\s*[-*+] (.*$)/gim, "<li>$1</li>");
  html = html.replace(/(<li>.*<\/li>)/gim, "<ul>$1</ul>");
  // Paragraphs
  html = html.replace(/\n\n/gim, "<br><br>");
  return html.trim();
}

function getCleanContent(content) {
  // Remove any temporary UI elements or formatting from the content
  return content.replace(/<[^>]*>?/gm, "").trim();
}

function copyToClipboard(content, event) {
  let htmlContent = content;
  ""
  if (!/<[a-z][\s\S]*>/i.test(content)) {
    htmlContent = markdownToHtml(content);
  }
  const cleanContent = getCleanContent(content);
  const button = event?.target?.closest?.("button");

  // Try to copy HTML if available
  if (navigator.clipboard?.write) {
    const htmlType = new ClipboardItem({
      "text/html": new Blob([htmlContent], { type: "text/html" }),
      "text/plain": new Blob([cleanContent], { type: "text/plain" }),
    });
    navigator.clipboard.write([htmlType]).then(
      () => showCopySuccess(button),
      () => fallbackCopy(cleanContent, button)
    );
  } else if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(cleanContent).then(
      () => showCopySuccess(button),
      () => fallbackCopy(cleanContent, button)
    );
  } else {
    fallbackCopy(cleanContent, button);
  }
}

function showCopySuccess(button) {
  if (!button) return;
  const originalHtml = button.innerHTML;
  button.innerHTML = '<i class="fas fa-check"></i> Copied!';
  setTimeout(() => {
    if (button) {
      button.innerHTML = originalHtml;
    }
  }, 2000);
}

function fallbackCopy(text, button) {
  let textarea;
  try {
    textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();

    if (document.execCommand("copy")) {
      showCopySuccess(button);
    } else {
      alert("Copy failed due to browser restrictions. Please copy manually.");
    }
  } catch (err) {
    console.error("Copy error:", err);
    alert("Copy failed. Please copy manually.");
  } finally {
    if (textarea && document.body.contains(textarea)) {
      document.body.removeChild(textarea);
    }
  }
}
