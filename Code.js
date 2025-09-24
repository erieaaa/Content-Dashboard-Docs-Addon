/**
 * @OnlyCurrentDoc
 * V42: Reliable Refresh Notification
 * - NEW: Implemented a robust visual notification system in the sidebar to prompt users to refresh after changing settings.
 * - MODIFIED: Sidebar.html now contains a notification banner and a red dot on the refresh icon that appear after settings are saved.
 * - REVERTED: showSettingsDialog() and saveTabSettings() are simplified. The complex server-side refresh logic has been removed in favor of a more reliable client-side approach.
 */

// --- Configuration ---
const CORE_TAGS_DEFAULT = [
  { name: 'intro', color: '#d9ead3' },
  { name: 'body', color: '#cfe2f3' },
  { name: 'conclusion', color: '#fce5cd' }
];

const DEFAULT_TAB_CONFIG = [
  { id: 'architectTab', name: 'Architect', defaultVisible: true },
  { id: 'taggerTab', name: 'Tagger', defaultVisible: true },
  { id: 'utilitiesTab', name: 'Utilities', defaultVisible: true }
];

/**
 * Creates and returns a card to be displayed as the add-on's homepage.
 * @param {Object} e The event object.
 * @return {CardService.Card} The homepage card.
 */
function onHomepage(e) {
  return createHomepageCard();
}

function createHomepageCard() {
  var openSidebarAction = CardService.newAction()
      .setFunctionName('showSidebar');

  var startButton = CardService.newTextButton()
      .setText('Start Dashboard')
      .setOnClickAction(openSidebarAction);

  var buttonSet = CardService.newButtonSet()
      .addButton(startButton);
      
  var section = CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("Click the button below to launch the main interface."))
      .addWidget(buttonSet); 
      
  var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Content Dashboard'))
      .addSection(section)
      .build();
      
  return card;
}

// --- Menu & Sidebar Creation ---
function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu() 
      .addItem('Start Content Dashboard', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar');
  const settingsResult = getTabSettings();
  html.tabs = settingsResult.success ? settingsResult.settings : DEFAULT_TAB_CONFIG.filter(t => t.defaultVisible);
  DocumentApp.getUi().showSidebar(html.evaluate().setTitle('Content Dashboard'));
}

function showOrganizerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('OrganizerDialog')
      .setWidth(900)
      .setHeight(600);
  DocumentApp.getUi().showModalDialog(html, 'Content Architect Board (Fullscreen)');
}

function showTagManagerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TagManagerDialog')
      .setWidth(500)
      .setHeight(550);
  DocumentApp.getUi().showModalDialog(html, 'Custom Tag Library');
}

// This function's only job is to display the modal dialog.
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
      .setWidth(400)
      .setHeight(450);
  DocumentApp.getUi().showModalDialog(html, 'Customize Sidebar Tabs');
}


// --- Master function for retrieving all tags ---
function getAllTags () {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    const allTagsJson = docProperties.getProperty('ALL_TAGS_ORDERED'); 
    if (allTagsJson) {
      return JSON.parse(allTagsJson);
    }
    else{
      docProperties.setProperty('ALL_TAGS_ORDERED', JSON.stringify(CORE_TAGS_DEFAULT));
      tagsOnlyView();
      return CORE_TAGS_DEFAULT;
    }
  } catch (e) {
     PropertiesService.getDocumentProperties().setProperty('ALL_TAGS_ORDERED', JSON.stringify(CORE_TAGS_DEFAULT));
    return CORE_TAGS_DEFAULT;
  }
}

function createNewTag(tagObject) {
  try {
    const tagName = tagObject.name.toLowerCase().trim().replace(/\s+/g, '-');
    const tagColor = tagObject.color.toLowerCase();
    const position = tagObject.position;
    if (!tagName) { return { success: false, message: "Tag name cannot be empty." }; }
    const allTags = getAllTags();
    if (allTags.some(t => t.name === tagName)) { return { success: false, message: `Tag "${tagName}" already exists.` }; }
    if (allTags.some(t => t.color === tagColor)) { return { success: false, message: `Color ${tagColor} is already in use.` }; }
    const newTag = { name: tagName, color: tagColor };
    if (position === 'end') { allTags.push(newTag); } 
    else {
      const targetIndex = allTags.findIndex(t => t.name === position);
      if (targetIndex !== -1) { allTags.splice(targetIndex + 1, 0, newTag); } 
      else { allTags.push(newTag); }
    }
     PropertiesService.getDocumentProperties().setProperty('ALL_TAGS_ORDERED', JSON.stringify(allTags));
  return { success: true, message: `Tag "${tagName}" created successfully.` };
  } catch (e) {
    return { success: false, message: "An unexpected error occurred while creating the tag." };
  }
}

// --- API Functions ---
function getInitialData() { 
  return getKanbanData(); 
}

function getKanbanData() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const paragraphs = body.getParagraphs();
    const kanbanData = [];
    const allTags = getAllTags(); 
    paragraphs.forEach((p, index) => {
      const text = p.getText();
      if (text.trim() === "") return;
      const match = text.match(/\[tag:\s*([\w-]+)-([\w\d.-]+)\s*\]$/);
      const cleanText = text.replace(/\s*\[tag:.*?\]$/, "");
      kanbanData.push({ 
        fullText: text,
        displayText: cleanText.split(' ').slice(0, 6).join(' ') + (cleanText.split(' ').length > 6 ? '...' : ''),
        originalIndex: index,
        category: match ? match[1].toLowerCase() : 'untagged',
        id: match ? match[2] : null
      });
    });
    return { kanbanData: kanbanData, allTags: allTags };
  } catch (e) {
    console.error(`Error in getKanbanData: ${e.message}`);
    return null;
  }
}

function applySmartTags(category) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    let maxId = 0;
    body.getParagraphs().forEach(p => {
      const text = p.getText();
      const tagRegex = new RegExp(`\\[tag:\\s*${category}-([\\w\\d.-]+)\\s*\\]`);
      const match = text.match(tagRegex);
      if (match) { 
          const currentId = parseInt(match[1]); 
          if (!isNaN(currentId) && currentId > maxId) { maxId = currentId; }
      }
    });
    let nextId = maxId + 1;
    const paragraphsToTag = new Set();
    const selection = doc.getSelection();
    const cursor = doc.getCursor();
    if (selection) {
      selection.getRangeElements().forEach(rangeElement => {
        let element = rangeElement.getElement();
        while (element && element.getParent) {
          if (element.getType() === DocumentApp.ElementType.PARAGRAPH || element.getType() === DocumentApp.ElementType.LIST_ITEM) {
            paragraphsToTag.add(element); break;
          }
          element = element.getParent();
        }
      });
    } else if (cursor) {
      let element = cursor.getElement();
      while (element && element.getParent) {
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH || element.getType() === DocumentApp.ElementType.LIST_ITEM) {
          paragraphsToTag.add(element); break;
        }
        element = element.getParent();
      }
    }
    if (paragraphsToTag.size === 0) return { success: false, message: "Please place your cursor or highlight text to tag." };
    const oldTagRegex = "\\s*\\[tag:\\s*[\\w-]+-[\\w\d.-]+\\s*\\]$";
    Array.from(paragraphsToTag).forEach(p => {
      p.asText().replaceText(oldTagRegex, "");
      p.asText().appendText(` [tag: ${category}-${nextId}]`);
      nextId++;
    });
    tagsOnlyView();
    return { success: true, message: `Successfully applied ${paragraphsToTag.size} "${category}" tag(s).` };
  } catch (e) {
    return { success: false, message: "An error occurred while applying tags." };
  }
}

function intelligentRenumberAllTags() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    let paragraphs = body.getParagraphs();
    let changesMade = false;
    for (let i = 1; i < paragraphs.length - 1; i++) {
      const prevMatch = paragraphs[i-1].getText().match(/\[tag:\s*([\w-]+)-[\w\d.-]+\s*\]$/);
      const currentText = paragraphs[i].getText();
      const currentMatch = currentText.match(/\[tag:\s*[\w-]+-[\w\d.-]+\s*\]$/);
      const nextMatch = paragraphs[i+1].getText().match(/\[tag:\s*([\w-]+)-[\w\d.-]+\s*\]$/);
      if (!currentMatch && prevMatch && nextMatch && prevMatch[1] === nextMatch[1] && currentText.trim() !== "") {
        const categoryToApply = prevMatch[1];
        paragraphs[i].appendText(` [tag: ${categoryToApply}-auto]`);
        changesMade = true;
      }
    }
    if (changesMade) { paragraphs = body.getParagraphs(); }
    const taggedParagraphs = paragraphs.filter(p => p.getText().match(/\[tag:\s*[\w-]+-[\w\d.-]+\s*\]$/));
    if (taggedParagraphs.length === 0) return { success: true, message: "No tags found to renumber."};
    const allTagNames = getAllTags().map(t => t.name);
    const counters = {};
    allTagNames.forEach(name => counters[name] = 1);
    const oldTagRegex = "\\s*\\[tag:\\s*[\\w-]+-[\\w\d.-]+\\s*\\]$";
    taggedParagraphs.forEach(p => {
      const match = p.getText().match(/\[tag:\s*([\w-]+)-[\w\d.-]+\s*\]$/);
      if (match) {
          const category = match[1].toLowerCase();
          if (counters.hasOwnProperty(category)) {
            p.asText().replaceText(oldTagRegex, "");
            p.asText().appendText(` [tag: ${category}-${counters[category]}]`);
            counters[category]++;
          }
      }
    });
    tagsOnlyView();
    return { success: true, message: "Tags discovered and renumbered." };
  } catch (e) {
    return { success: false, message: "An error occurred during renumbering." };
  }
}

function reorganizeFromKanban(kanbanColumns) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const paragraphsToDelete = new Set();
    const paragraphsToInsert = [];
    const allOriginalParas = body.getParagraphs();
    const allTagNames = getAllTags().map(t => t.name);
    Object.values(kanbanColumns).flat().forEach(item => {
        if (item && typeof item.originalIndex !== 'undefined') {
          const p = allOriginalParas[item.originalIndex];
          if (p) paragraphsToDelete.add(p);
        }
    });
    allTagNames.forEach(category => {
      const cards = kanbanColumns[category] || [];
      cards.sort((a, b) => {
          const idA = a.id ? a.id.toString() : '';
          const idB = b.id ? b.id.toString() : '';
          return idA.localeCompare(idB, undefined, {numeric: true});
      });
      cards.forEach(card => {
          if (card && typeof card.originalIndex !== 'undefined') {
              const originalPara = allOriginalParas[card.originalIndex];
              if (originalPara) {
                  paragraphsToInsert.push(originalPara.copy());
              }
          }
      });
    });
    if (paragraphsToDelete.size === 0) return { success: false, message: "No paragraphs found to reorganize." };
    const deletionIndices = Array.from(paragraphsToDelete).map(p => body.getChildIndex(p)).sort((a,b) => b-a);
    deletionIndices.forEach(index => {
        if (body.getNumChildren() > 1) { 
            body.getChild(index).removeFromParent(); 
        } else { 
            body.getChild(index).asParagraph().clear(); 
        }
    });
    paragraphsToInsert.forEach((para, index) => { 
        body.insertParagraph(index, para); 
    });
    intelligentRenumberAllTags();
    return { success: true, message: "Document has been rebuilt!" };
  } catch (e) {
    return { success: false, message: "An error occurred while rebuilding the document." };
  }
}

// --- View Mode Functions ---
function standardView() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    body.getParagraphs().forEach(p => {
      p.setBackgroundColor(null);
      const text = p.getText();
      const match = text.match(/(\[tag:\s*[\w-]+-[\w\d.-]+\s*\])$/);
      if (match) {
        const tagText = match[1];
        const startIndex = text.lastIndexOf(tagText);
        const endIndex = startIndex + tagText.length - 1;
        const textElement = p.editAsText();
        textElement.setFontSize(startIndex, endIndex, 1);
        textElement.setForegroundColor(startIndex, endIndex, '#ffffff');
      }
    });
    return { success: true, message: 'Switched to Standard View.' };
  } catch (e) {
    return { success: false, message: "Could not apply Standard View." };
  }
}

function tagsOnlyView() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const allTags = getAllTags();
    const tagColorMap = allTags.reduce((map, tag) => {
        map[tag.name] = tag.color;
        return map;
    }, {});
    body.getParagraphs().forEach(p => {
      p.setBackgroundColor(null);
      p.editAsText().setFontSize(null).setForegroundColor(null);
      const text = p.getText();
      const match = text.match(/(\[tag:\s*([\w-]+)-[\w\d.-]+\s*\])$/);
      if (match) {
        const tagText = match[1];
        const category = match[2].toLowerCase();
        const color = tagColorMap[category];
        if (color) {
          const startIndex = text.lastIndexOf(tagText);
          const endIndex = startIndex + tagText.length - 1;
          p.editAsText().setBackgroundColor(startIndex, endIndex, color);
        }
      }
    });
    return { success: true, message: 'Switched to Tags Only View.' };
  } catch (e) {
    return { success: false, message: "Could not apply Tags Only View." };
  }
}

function structureAuditView() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const allTags = getAllTags();
    const tagColorMap = allTags.reduce((map, tag) => {
      map[tag.name] = tag.color;
      return map;
    }, {});
    body.getParagraphs().forEach(p => {
      const text = p.getText();
      const match = text.match(/(\[tag:\s*[\w-]+-[\w\d.-]+\s*\])$/);
      if (match) {
        const category = match[2].toLowerCase();
        const color = tagColorMap[category];
        if (color) {
          p.setBackgroundColor(color);
          const tagText = match[1];
          const startIndex = text.lastIndexOf(tagText);
          const endIndex = startIndex + tagText.length - 1;
          p.editAsText().setFontSize(startIndex, endIndex, 1);
          p.editAsText().setForegroundColor(startIndex, endIndex, '#ffffff');
        }
      } else {
        p.setBackgroundColor(null);
      }
    });
    return { success: true, message: 'Switched to Structure Audit View.' };
  } catch (e) {
    return { success: false, message: "Could not apply Structure Audit View." };
  }
}

// --- Tag Management Functions ---
function updateTag(oldName, newTagObject) {
  try {
    const newName = newTagObject.name.toLowerCase().trim().replace(/\s+/g, '-');
    const newColor = newTagObject.color.toLowerCase();
    if (!newName) { return { success: false, message: "Tag name cannot be empty." }; }
    const allTags = getAllTags();
    const tagIndex = allTags.findIndex(t => t.name === oldName);
    if (tagIndex === -1) { return { success: false, message: `Original tag "${oldName}" not found.` }; }
    if (allTags.some((t, i) => i !== tagIndex && t.name === newName)) { return { success: false, message: `Tag name "${newName}" is already in use.` }; }
    if (allTags.some((t, i) => i !== tagIndex && t.color === newColor)) { return { success: false, message: `Color ${newColor} is already in use.` }; }
    allTags[tagIndex] = { name: newName, color: newColor };
  PropertiesService.getDocumentProperties().setProperty('ALL_TAGS_ORDERED', JSON.stringify(allTags));
    if (oldName !== newName) {
      const body = DocumentApp.getActiveDocument().getBody();
      const searchPattern = `\\[tag:\\s*${oldName}-([\\w\\d.-]+)\\s*\\]`;
      const replacement = `[tag: ${newName}-$1]`;
      body.replaceText(searchPattern, replacement);
    }
    tagsOnlyView();
    return { success: true, message: "Tag updated successfully." };
  } catch (e) {
    return { success: false, message: "An unexpected error occurred while updating the tag." };
  }
}

function deleteTag(tagName) {
  try {
    const allTags = getAllTags();
    const updatedTags = allTags.filter(t => t.name !== tagName);
    if (allTags.length === updatedTags.length) { return { success: false, message: `Tag "${tagName}" not found.` }; }
    PropertiesService.getDocumentProperties().setProperty('ALL_TAGS_ORDERED', JSON.stringify(updatedTags));
    const body = DocumentApp.getActiveDocument().getBody();
    const searchPattern = `\\[tag:\\s*${tagName}-[\\w\\d.-]+\\s*\\]$`;
    body.replaceText(searchPattern, "");
    tagsOnlyView();
    return { success: true, message: `Tag "${tagName}" deleted.` };
  } catch (e) {
    return { success: false, message: "An unexpected error occurred while deleting the tag." };
  }
}

// --- Tab Settings Functions ---
function getTabSettings() {
  try {
    const docProperties = PropertiesService.getDocumentProperties();
    const settingsJson = docProperties.getProperty('SIDEBAR_TAB_SETTINGS');
    if (settingsJson) {
      return { success: true, settings: JSON.parse(settingsJson), allTabs: DEFAULT_TAB_CONFIG };
    } else {
      const defaultSettings = DEFAULT_TAB_CONFIG.filter(tab => tab.defaultVisible);
      docProperties.setProperty('SIDEBAR_TAB_SETTINGS', JSON.stringify(defaultSettings));
      return { success: true, settings: defaultSettings, allTabs: DEFAULT_TAB_CONFIG };
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// This function's only job is to validate and save the settings.
function saveTabSettings(settings) {
  try {
    if (!Array.isArray(settings)) {
      throw new Error("Invalid settings format. Expected an array.");
    }
    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.setProperty('SIDEBAR_TAB_SETTINGS', JSON.stringify(settings));
    return { success: true, message: "Settings saved!", settings: settings }; 
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- Goal Tracker Functions ---

/**
 * Retrieves goal settings and current document progress.
 * Resets the daily word count if it's a new day.
 */
function getGoalSettings() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const settingsJson = userProperties.getProperty('GOAL_SETTINGS');
    const progressJson = userProperties.getProperty('GOAL_PROGRESS');
    const customGoalsJson = userProperties.getProperty('CUSTOM_GOALS');

    const settings = settingsJson ? JSON.parse(settingsJson) : { target: 5000, dueDate: '', dailyTarget: 500 };
    let progress = progressJson ? JSON.parse(progressJson) : { startOfDayCount: 0, date: '' };
    const customGoals = customGoalsJson ? JSON.parse(customGoalsJson) : []; // Correctly initialize customGoals

    const currentWordCount = getWordCount();
    const today = new Date().toLocaleDateString();

    // If the saved date is not today, it's a new day. Reset the daily counter.
    if (progress.date !== today) {
      progress.date = today;
      progress.startOfDayCount = currentWordCount;
      userProperties.setProperty('GOAL_PROGRESS', JSON.stringify(progress));
    }
    
    const wordsWrittenToday = Math.max(0, currentWordCount - progress.startOfDayCount);

    return { 
      success: true, 
      settings: settings,
      currentWordCount: currentWordCount,
      wordsWrittenToday: wordsWrittenToday,
      customGoals: customGoals // Pass the customGoals array
    };
  } catch (e) {
    return { success: false, message: "Could not load goal settings: " + e.message };
  }
}

/**
 * Saves the user's goal settings.
 */
function saveGoalSettings(settings) {
  try {
    // Basic validation
    if (typeof settings.target !== 'number' || typeof settings.dailyTarget !== 'number' || typeof settings.dueDate !== 'string') {
      throw new Error("Invalid settings object received.");
    }
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('GOAL_SETTINGS', JSON.stringify(settings));
    // After saving, return the updated state
    return getGoalSettings(); 
  } catch (e) {
    return { success: false, message: "Error saving settings: " + e.message };
  }
}

/**
 * Helper function to count words in the document body.
 */
function getWordCount() {
    const text = DocumentApp.getActiveDocument().getBody().getText();
    // Use a regex to split by whitespace and filter out empty strings
    const words = text.split(/\s+/).filter(word => word.length > 0);
    return words.length;
}

/**
 * Saves the user's list of custom milestone goals.
 */
function saveCustomGoals(goals) {
  try {
    if (!Array.isArray(goals)) {
      throw new Error("Invalid data format for custom goals.");
    }
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('CUSTOM_GOALS', JSON.stringify(goals));
    return { success: true, message: "Milestones saved." };
  } catch (e) {
    return { success: false, message: "Error saving milestones: " + e.message };
  }
}

/**
 * Shows the dialog for managing project milestones.
 */
function showMilestonesDialog() {
  const html = HtmlService.createHtmlOutputFromFile('MilestonesDialog')
      .setWidth(600)
      .setHeight(500);
  DocumentApp.getUi().showModalDialog(html, 'Manage Project Milestones');
}