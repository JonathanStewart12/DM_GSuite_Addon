// TODO - [2] verify that DM System URLs must first be whitelisted in the addon manifest
// if so, ensure this is recorded in the appropriate documentation

let MimeTypes = {
    "gdoc": "application/vnd.google-apps.document",
    "sheets": "application/vnd.google-apps.spreadsheet",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf"
};
let FileExtensions = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/pdf": ".pdf"
};

function onLoadDocs(eventData) {
    Context.eventData = eventData;
    let factory = new DocsAppCardFactory();
    let card = factory.create();
    return card.build();
}

function onLoadSheets(eventData) {
    Context.eventData = eventData;
    let factory = new SheetsAppCardFactory();
    let card = factory.create();
    return card.build();
}

function onLoadDrive(eventData) {
    Context.eventData = eventData;
    Logger.log(JSON.stringify(Context.eventData.drive.selectedItems));
    let factory = new DriveAppCardFactory();
    let card = factory.create();
    return card.build();
}

function onLoadGmail(eventData) {
    Context.eventData = eventData;
    let factory : GmailAppCardFactory;
    Logger.log(JSON.stringify(eventData));
    if(eventData.hasOwnProperty("formInput") && eventData.formInput.hasOwnProperty("saveEmailOption")) {
        factory = new GmailAppCardFactory(eventData.formInput.saveEmailOption);
    } else {
        factory = new GmailAppCardFactory();
    }
    let card = factory.create();
    return card.build();
}

function onClickClearPropertiesButton(eventData) {
    PropertiesService.getScriptProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popToRoot())
        .build();
}

function onChangeDmSystemDropdown(eventData) {
    Context.eventData = eventData;
    let id = "";
    if(eventData.formInput.hasOwnProperty("dmSystem")) {
        id = eventData.formInput.dmSystem;
    }
    DMSystemRepository.setSelectedDMSystemId(id);
}

function onClickNewDmSystemButton(eventData) {
    Context.eventData = eventData;
    let factory = new DMSystemCardFactory(true);
    return factory.create().build();
}

function onClickManageDmSystemButton(eventData) {
    Context.eventData = eventData;
    if(!eventData.formInput.hasOwnProperty("dmSystem")) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must select a DM System from the dropdown."))
            .build();
    }
    let factory = new DMSystemCardFactory(false, eventData.formInput.dmSystem);
    return factory.create().build();
}

function onClickDeleteDmSystemButton(eventData) {
    Context.eventData = eventData;
    let dmSystemId : string;
    if(eventData.hasOwnProperty("parameters") && eventData.parameters.hasOwnProperty("dmSystem")) {
        dmSystemId = eventData.parameters.dmSystem;
    } else if (eventData.formInput.hasOwnProperty("dmSystem")) {
        dmSystemId = eventData.formInput.dmSystem;
    } else {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("Error deleting DM System. Please try again."))
            .build();
    }
    let feedbackMessage : string;
    if(DMSystemRepository.delete(dmSystemId)) {
        feedbackMessage = "Deleted DM System.";
    } else {
        feedbackMessage = "Error deleting DM System.";
    }
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(feedbackMessage))
        .setNavigation(CardService.newNavigation().popToRoot().updateCard(getRootCard(eventData).build()))
        .build();
}

function onClickDmSystemSaveButton(eventData) {
    Context.eventData = eventData;
    Logger.log(JSON.stringify(eventData));
    if(!eventData.formInput.hasOwnProperty("name")) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must enter a name."))
            .build();
    }
    if(!eventData.formInput.hasOwnProperty("url")) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must enter a URL."))
            .build();
    }
    let dmSystemId : string;
    switch(eventData.parameters.mode) {
        case "create":
            dmSystemId = DMSystemRepository.create(eventData.formInput.name, eventData.formInput.url);
            break;
        case "update":
            let dmSystem = new DMSystem();
            dmSystem.id = eventData.parameters.id;
            dmSystem.name = eventData.formInput.name;
            dmSystem.url = eventData.formInput.url;
            DMSystemRepository.update(dmSystem);
            dmSystemId = dmSystem.id;
            break;
    }
    let card = new DMSystemCardFactory(false, dmSystemId).create();
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("DM System " + (eventData.parameters.mode == "create" ? "created" : "updated") + "."))
        .setNavigation(CardService.newNavigation().updateCard(card.build()))
        .build();
}

function getRootCard(eventData) {
    Context.eventData = eventData;
    switch(eventData.hostApp) {
        case "docs":
            return new DocsAppCardFactory().create();
        case "drive":
            return new DriveAppCardFactory().create();
        case "gmail":
            return new GmailAppCardFactory().create();
        case "sheets":
            return new SheetsAppCardFactory().create();
    }
    return null;
}

function onClickDmSystemReturnButton(eventData) {
    Context.eventData = eventData;
    let rootCard : GoogleAppsScript.Card_Service.CardBuilder;
    switch(eventData.hostApp) {
        case "docs":
            rootCard = new DocsAppCardFactory().create();
            break;
        case "drive":
            rootCard = new DriveAppCardFactory().create();
            break;
        case "gmail":
            rootCard = new GmailAppCardFactory().create();
            break; 
        case "sheets":
            rootCard = new SheetsAppCardFactory().create();
            break;
    }
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popToRoot().updateCard(rootCard.build()))
        .build();
}

function onClickSaveDocument(eventData) {
    Context.eventData = eventData;
    let dmSystem : DMSystem = DMSystemRepository.getSelectedDMSystem();
    if(dmSystem == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must select a DM System from the dropdown."))
            .build();
    } else {
        return FileHelper.saveToDocumentManager([Context.eventData.docs.id], dmSystem);
    }
}

function onClickSaveSpreadsheet(eventData) {
    Context.eventData = eventData;
    let dmSystem : DMSystem = DMSystemRepository.getSelectedDMSystem();
    if(dmSystem == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must select a DM System from the dropdown."))
            .build();
    } else {
        return FileHelper.saveToDocumentManager([Context.eventData.sheets.id], dmSystem);
    }
}

function onClickSaveDrive(eventData) {
    Context.eventData = eventData;
    let dmSystem : DMSystem = DMSystemRepository.getSelectedDMSystem();
    if(dmSystem == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must select a DM System from the dropdown."))
            .build();
    } else {
        let fileIds : string[] = new Array();
        for(let i=0; i<eventData.drive.selectedItems.length; i++) {
            fileIds.push(eventData.drive.selectedItems[i].id);
        }
        return FileHelper.saveToDocumentManager(fileIds, dmSystem);
    }
}

function onClickSaveEmail(eventData) {
    Context.eventData = eventData;
    let dmSystem : DMSystem = DMSystemRepository.getSelectedDMSystem();
    if(dmSystem == null) {
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("You must select a DM System from the dropdown."))
            .build();
    } else {
        for(let blob of getBlobsToSave()) {
            UrlFetchApp.fetch(dmSystem.getAnonymousUploadURL(), DMSystem.getAnonymousUploadOptions(blob));
        }
        dmSystem.addSavedFileId(eventData.gmail.messageId);
        DMSystemRepository.update(dmSystem);
        return DMSystem.createRedirectToImportPageActionResponse(dmSystem.getImportPageURL());
    }

    function getBlobsToSave() {
        let result = new Array();
        if (eventData.formInput.saveEmailOption == "merge") {
            result.push(getEmailBlob());
        } else if (eventData.formInput.saveEmailOption == "attachments") {
            result = getAttachmentBlobs();
        }
        return result;
    }

    function getEmailBlob() {
        let emailBlob = Utilities.newBlob(Utilities.base64DecodeWebSafe(getEmailData(eventData.gmail.messageId)), "text/plain");
        emailBlob.setName(GmailApp.getMessageById(eventData.gmail.messageId).getSubject() + ".eml");
        return emailBlob;
    }

    function getEmailData(emailId: any) {
        let idToken = ScriptApp.getIdentityToken();
        let body = idToken.split('.')[1];
        let decoded = Utilities.newBlob(Utilities.base64Decode(body)).getDataAsString();
        let payload = JSON.parse(decoded);
        let userId = payload.sub;
        let url = `https://gmail.googleapis.com/gmail/v1/users/${userId}/messages/${emailId}?format=raw`;
        let httpResponse = UrlFetchApp.fetch(url, {
            method: "get",
            headers: {
                Authorization: "Bearer " + ScriptApp.getOAuthToken()
            }
        });
        let emailDataRaw = JSON.parse(httpResponse.getContentText()).raw;
        return emailDataRaw;
    }

    function getAttachmentBlobs() {
        let result = new Array();
        let gmailMessage = GmailApp.getMessageById(eventData.gmail.messageId);
        let selectedAttachmentIds = eventData.formInputs.saveAttachments;
        for (let attachment of gmailMessage.getAttachments()) {
            if (selectedAttachmentIds.includes(attachment.getHash())) {
                result.push(attachment.copyBlob());
            }
        }
        return result;
    }
}

function onAuthoriseFileScope(eventData) {
    // @ts-ignore
    return CardService.newEditorFileScopeActionResponseBuilder()
        .requestFileScopeForActiveDocument()
        .build();
}

class CommonCardFactory {
    public static create() : GoogleAppsScript.Card_Service.CardBuilder {
        let result = CardService.newCardBuilder();
        result.addSection(this.createDmSystemHeader());
        return result;
    }

    public static createNotificationSection(message : string) : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        let textParagraph = CardService.newTextParagraph();
        textParagraph.setText(message);
        result.addWidget(textParagraph);
        return result;
    }

    private static createDmSystemHeader() : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        result.addWidget(this.createDmSystemDropdown());
        result.addWidget(this.createDmSystemButtonSet());
        result.addWidget(this.createClearPropertiesButton());
        return result;
    }

    private static createDmSystemDropdown() : GoogleAppsScript.Card_Service.SelectionInput {
        let result = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("DM System")
            .setFieldName("dmSystem")
            .setOnChangeAction(CardService.newAction().setFunctionName("onChangeDmSystemDropdown"));
        this.populateDropdown(result);
        return result;
    }

    private static populateDropdown(dropdown: GoogleAppsScript.Card_Service.SelectionInput) {
        let dmSystems = DMSystemRepository.getDmSystems();
        dropdown.addItem("", "", dmSystems.length === 0);
        for(let i=0; i<dmSystems.length; i++) {
            let dmSystem = dmSystems[i];
            dropdown.addItem(dmSystem.name, dmSystem.id, dmSystem.id === DMSystemRepository.getSelectedDMSystemId());
        }
    }

    private static createDmSystemButtonSet(): GoogleAppsScript.Card_Service.ButtonSet {
        let result = CardService.newButtonSet();
        result.addButton(this.createNewButton());
        result.addButton(this.createManageButton());
        result.addButton(this.createDeleteButton());
        return result;
    }

    private static createNewButton(): GoogleAppsScript.Card_Service.Button {
        let result = CardService.newTextButton()
            .setText("New")
            .setOnClickAction(CardService.newAction().setFunctionName("onClickNewDmSystemButton").setParameters({"mode": "create"}))
        return result;
    }

    private static createManageButton(): GoogleAppsScript.Card_Service.Button {
        let result = CardService.newTextButton()
            .setText("Manage")
            .setOnClickAction(CardService.newAction().setFunctionName("onClickManageDmSystemButton"));
        return result;
    }

    private static createDeleteButton(): GoogleAppsScript.Card_Service.Button {
        let result = CardService.newTextButton()
            .setText("Delete")
            .setOnClickAction(CardService.newAction().setFunctionName("onClickDeleteDmSystemButton"));
        return result;
    }

    private static createClearPropertiesButton(): GoogleAppsScript.Card_Service.Widget {
        let action = CardService.newAction()
            .setFunctionName("onClickClearPropertiesButton");
        return CardService.newTextButton()
            .setText("Clear Properties")
            .setOnClickAction(action);
    }

}

class DMSystem {
    public batchId: string;
    public url: string;
    public name: string;
    public id: string;
    public savedFileIds : string[];

    constructor() {
        this.batchId = GuidCreator.create();
    }

    public getAnonymousUploadURL() : string {
        return this.url + "/Import/AnonymousUpload/?batchId=" + this.batchId + "&source=GSuite";
    }

    public getImportPageURL() : string {
        return this.url + "/Import/ImportBatch/?batchId=" + this.batchId;
    }

    public addSavedFileId(id : string) {
        if(typeof this.savedFileIds === "undefined") {
            this.savedFileIds = new Array();
        }
        this.savedFileIds.push(id);
    }

    public static getAnonymousUploadOptions(blob : GoogleAppsScript.Base.Blob) : any {
        return {
            method: "post",
            payload: {
                file: blob
            }
        };
    }

    public static createRedirectToImportPageActionResponse(importPageURL : string) : GoogleAppsScript.Card_Service.ActionResponse {
        var openLink = CardService.newOpenLink();
        openLink.setUrl(importPageURL);
        var actionResponse = CardService.newActionResponseBuilder();
        actionResponse.setOpenLink(openLink);
        return actionResponse.build();
    }

}

class GuidCreator {
    public static create() : string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }
}

class DMSystemRepository {
    public static getSelectedDMSystem() : DMSystem {
        return DMSystemRepository.getDmSystem(DMSystemRepository.getSelectedDMSystemId());
    }

    public static setSelectedDMSystemId(id : string) {
        PropertiesService.getUserProperties().setProperty("selectedDMSystem", id);
    }

    public static getSelectedDMSystemId() : string {
        return PropertiesService.getUserProperties().getProperty("selectedDMSystem");
    }

    public static getDmSystem(id : string) : DMSystem {
        let dmSystems = DMSystemRepository.getDmSystems();
        for(let i=0; i<dmSystems.length; i++) {
            let dmSystem = dmSystems[i];
            if(id === dmSystem.id) {
                return dmSystem;
            }
        }
        return null;
    }

    public static getDmSystems() : DMSystem[] {
        let propertyValue = PropertiesService.getScriptProperties().getProperty("dmSystems");
        let result = new Array();
        if(propertyValue) {
            let array = JSON.parse(propertyValue);
            array.forEach(function(element) {
                result.push(DMSystemRepository.parse(element));
            });
        }
        return result;
    }

    public static parse(data) : DMSystem {
        return Object.assign(new DMSystem(), data);
    }

    // TODO - [3] Prevent duplicate environments from being created. Use the URL value
    // to determine this.
    public static create(name : string, url : string) : string {
        let dmSystems : DMSystem[] = DMSystemRepository.getDmSystems();
        let newDmSystem = new DMSystem();
        newDmSystem.name = name;
        newDmSystem.url = url; 
        newDmSystem.id = GuidCreator.create();
        dmSystems.push(newDmSystem);
        PropertiesService.getScriptProperties().setProperty("dmSystems", JSON.stringify(dmSystems));
        return newDmSystem.id;
    }

    public static delete(id : string) : boolean {
        let dmSystems : DMSystem[] = DMSystemRepository.getDmSystems();
        let index = -1;
        for(let i=0; i<dmSystems.length; i++) {
            if(dmSystems[i].id === id) {
                index = i;
                break;
            }
        }
        if(index > -1) {
            dmSystems.splice(index, 1);
            PropertiesService.getScriptProperties().setProperty("dmSystems", JSON.stringify(dmSystems));
            return true;
        }
        return false;
    }

    public static update(dmSystemToUpdate : DMSystem) {
        let dmSystems : DMSystem[] = DMSystemRepository.getDmSystems();
        for(let i=0; i<dmSystems.length; i++) {
            if(dmSystems[i].id == dmSystemToUpdate.id) {
                dmSystems[i] = dmSystemToUpdate;
                break;
            }
        }
        PropertiesService.getScriptProperties().setProperty("dmSystems", JSON.stringify(dmSystems));
    }

}

class DMSystemCardFactory {
    private isCreateMode: boolean;
    private dmSystem : DMSystem;

    constructor(isCreateMode : boolean, dmSystemId? : string) {
        this.isCreateMode = isCreateMode;
        if(!isCreateMode && typeof dmSystemId !== "undefined") {
            this.dmSystem = DMSystemRepository.getDmSystem(dmSystemId);
        }
    }

    public create() : GoogleAppsScript.Card_Service.CardBuilder {
        return CardService.newCardBuilder()
            .addSection(CardService.newCardSection()
            .addWidget(this.createNameField())
            .addWidget(this.createURLField())
            .addWidget(this.createButtonSet()));
    }

    private createNameField(): GoogleAppsScript.Card_Service.TextInput {
        let result = CardService.newTextInput()
            .setFieldName("name")
            .setTitle("Name");
        if(!this.isCreateMode && this.dmSystem) {
            result.setValue(this.dmSystem.name)
        }
        return result;
    }

    private createURLField(): GoogleAppsScript.Card_Service.TextInput {
        let result = CardService.newTextInput()
            .setFieldName("url")
            .setTitle("URL");
        if(!this.isCreateMode && this.dmSystem) {
            result.setValue(this.dmSystem.url)
        }
        return result;
    }

    private createButtonSet() : GoogleAppsScript.Card_Service.ButtonSet {
        let result = CardService.newButtonSet();
        result.addButton(this.createSaveButton());
        if(!this.isCreateMode && (typeof this.dmSystem !== "undefined" && this.dmSystem != null)) {
            result.addButton(this.createDeleteButton());
        }
        result.addButton(this.createReturnButton());
        return result;
    }

    private createSaveButton() : GoogleAppsScript.Card_Service.TextButton {
        let action = CardService.newAction()
            .setFunctionName("onClickDmSystemSaveButton")
            .setParameters(this.createSaveButtonParameters());
        return CardService.newTextButton()
            .setText("Save")
            .setOnClickAction(action);
    }

    private createSaveButtonParameters() {
        Logger.log("creating save button parameters");
        let result = {};
        result["mode"] = this.isCreateMode ? "create" : "update";
        Logger.log("mode: " + result["mode"]);
        if(!this.isCreateMode && (typeof this.dmSystem !== "undefined" && this.dmSystem != null)) {
            result["id"] = this.dmSystem.id;
        }
        return result;
    }

    private createReturnButton() : GoogleAppsScript.Card_Service.TextButton {
        let action = CardService.newAction()
            .setFunctionName("onClickDmSystemReturnButton");
        return CardService.newTextButton()
            .setText("Return")
            .setOnClickAction(action);
    }

    private createDeleteButton(): GoogleAppsScript.Card_Service.Button {
        let action = CardService.newAction()
            .setFunctionName("onClickDeleteDmSystemButton")
            .setParameters({
                "dmSystem": this.dmSystem.id
            });
        return CardService.newTextButton()
            .setText("Delete")
            .setOnClickAction(action);
    }
}

class DocsAppCardFactory {
    private card : GoogleAppsScript.Card_Service.CardBuilder;

    constructor() {
        this.card = CommonCardFactory.create();
    }

    create() : GoogleAppsScript.Card_Service.CardBuilder {
        if(this.isFileAlreadySaved()) {
            this.card.addSection(CommonCardFactory.createNotificationSection("This document is already saved."));
        }
        this.card.addSection(this.createSection());
        return this.card;
    }

    private isFileAlreadySaved() {
        let dmSystem = DMSystemRepository.getSelectedDMSystem();
        return dmSystem != null && typeof dmSystem.savedFileIds !== "undefined" && dmSystem.savedFileIds.includes(Context.eventData.docs.id);
    }

    createSection() : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        if(this.addonHasFileScopePermission()) {
            if(FileHelper.isFileEmpty(Context.eventData.docs.id)) {
                return CommonCardFactory.createNotificationSection("Please add content to the document and re-load the addon.");
            } else {
                this.addSaveDocWidgets(result);
            }
        } else {
            this.addAuthoriseWidgets(result);
        }
        return result;
    }

    private addonHasFileScopePermission() : Boolean {
        return Context.eventData.docs.addonHasFileScopePermission;
    }

    private addSaveDocWidgets(section: GoogleAppsScript.Card_Service.CardSection) {
        if(FileHelper.isGoogleType(Context.eventData.docs.id)) {
            section.addWidget(this.createFileTypeDropdown());
        }
        section.addWidget(this.createSaveButton());
        section.setHeader("Save Document");
    }

    private addAuthoriseWidgets(section: GoogleAppsScript.Card_Service.CardSection) {
        let message = CardService.newTextParagraph()
            .setText("Authorisation is required to access this document. Please click the Authorise button below.")
        let authorise = CardService.newTextButton()
            .setText("Authorise")
            .setOnClickAction(CardService.newAction().setFunctionName("onAuthoriseFileScope"));
        section.addWidget(message);
        section.addWidget(authorise);
    }

    private createFileTypeDropdown(): GoogleAppsScript.Card_Service.Widget {
        return CardService.newSelectionInput()
            .setTitle("File Type")
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setFieldName("fileType")
            .addItem("Word", MimeTypes.docx, true)
            .addItem("PDF", MimeTypes.pdf, false)
    }

    private createSaveButton(): GoogleAppsScript.Card_Service.Widget {
        return CardService.newTextButton()
            .setText("Save")
            .setOnClickAction(CardService.newAction().setFunctionName("onClickSaveDocument"));
    }
}

class SheetsAppCardFactory {
    private card : GoogleAppsScript.Card_Service.CardBuilder;

    constructor() {
        this.card = CommonCardFactory.create();
    }

    create() : GoogleAppsScript.Card_Service.CardBuilder {
        if(this.isFileAlreadySaved()) {
            this.card.addSection(CommonCardFactory.createNotificationSection("This spreadsheet is already saved."));
        }
        this.card.addSection(this.createSection());
        return this.card;
    }

    private isFileAlreadySaved() {
        let dmSystem = DMSystemRepository.getSelectedDMSystem();
        return dmSystem != null && typeof dmSystem.savedFileIds !== "undefined" && dmSystem.savedFileIds.includes(Context.eventData.sheets.id);
    }

    createSection() : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        if(this.addonHasFileScopePermission()) {
            if(FileHelper.isFileEmpty(Context.eventData.sheets.id)) {
                return CommonCardFactory.createNotificationSection("Please add content to the spreadsheet and re-load the addon.");
            } else {
                this.addSaveSpreadsheetWidgets(result);
            }
        } else {
            this.addAuthoriseWidgets(result);
        }
        return result;
    }

    private addonHasFileScopePermission() : Boolean {
        return Context.eventData.sheets.addonHasFileScopePermission;
    }

    private addSaveSpreadsheetWidgets(section: GoogleAppsScript.Card_Service.CardSection) {
        if(FileHelper.isGoogleType(Context.eventData.sheets.id)) {
            section.addWidget(this.createFileTypeDropdown());
        }
        section.addWidget(this.createSaveButton());
        section.setHeader("Save Spreadsheet");
    }

    private addAuthoriseWidgets(section: GoogleAppsScript.Card_Service.CardSection) {
        let message = CardService.newTextParagraph()
            .setText("Authorisation is required to access this spreadsheet. Please click the Authorise button below.")
        let authorise = CardService.newTextButton()
            .setText("Authorise")
            .setOnClickAction(CardService.newAction().setFunctionName("onAuthoriseFileScope"));
        section.addWidget(message);
        section.addWidget(authorise);
    }

    private createFileTypeDropdown(): GoogleAppsScript.Card_Service.Widget {
        return CardService.newSelectionInput()
            .setTitle("File Type")
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setFieldName("fileType")
            .addItem("Excel", MimeTypes.xlsx, true)
            .addItem("PDF", MimeTypes.pdf, false)
    }

    private createSaveButton(): GoogleAppsScript.Card_Service.Widget {
        return CardService.newTextButton()
            .setText("Save")
            .setOnClickAction(CardService.newAction().setFunctionName("onClickSaveSpreadsheet"));
    }
}

class DriveAppCardFactory {
    private card : GoogleAppsScript.Card_Service.CardBuilder;

    constructor() {
        this.card = CommonCardFactory.create();
    }

    create() : GoogleAppsScript.Card_Service.CardBuilder {
        if(Context.eventData.drive.hasOwnProperty("selectedItems") && this.selectionContainsSavedFiles()) {
            this.card.addSection(this.createSaveStatusNotificationSection());
        }
        this.card.addSection(this.createSection())
        return this.card;
    }

    private selectionContainsSavedFiles() : boolean {
        let dmSystem = DMSystemRepository.getSelectedDMSystem();
        if(dmSystem == null || typeof dmSystem.savedFileIds === "undefined") {
            return false;
        }
        for(let selectedItem of Context.eventData.drive.selectedItems) {
            if(dmSystem.savedFileIds.includes(selectedItem.id)) {
                return true;
            }
        }
        return false;
    }

    private createSaveStatusNotificationSection(): GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        let textArray = new Array();
        if(Context.eventData.drive.selectedItems.length == 1) {
            textArray.push("The selected file is already saved.");
        } else if (Context.eventData.drive.selectedItems.length > 1) {
            textArray.push("The files listed below are already saved to Document Manager.<br />");
            let dmSystem = DMSystemRepository.getSelectedDMSystem();
            for(let selectedItem of Context.eventData.drive.selectedItems) {
                if(dmSystem.savedFileIds.includes(selectedItem.id)) {
                    textArray.push(`<li>${selectedItem.title}</li>`);
                }
            }
            textArray.push("<br />");
        }
        result.addWidget(CardService.newTextParagraph().setText(textArray.join("<br />")));
        return result;
    }

    createSection() : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        let fileCount = this.getFileCount();
        if(fileCount == 0) {
            result.addWidget(CardService.newTextParagraph().setText("Please make a selection."));
        } else {
            if(fileCount == 1) {
                this.buildUIForSingleSelection(result);
            } 
            result.addWidget(CardService.newTextButton().setText("Save").setOnClickAction(CardService.newAction().setFunctionName("onClickSaveDrive")))
        }
        return result;
    }

    private buildUIForSingleSelection(result: GoogleAppsScript.Card_Service.CardSection) {
        let selectedItem = Context.eventData.drive.selectedItems[0];
        if (FileHelper.isGoogleType(selectedItem.id)) {
            let fileTypeDropdown = CardService.newSelectionInput()
                .setType(CardService.SelectionInputType.DROPDOWN)
                .setFieldName("fileType")
                .setTitle("File Type")
                .addItem("PDF", MimeTypes.pdf, false);
            if (selectedItem.mimeType == MimeTypes.gdoc) {
                fileTypeDropdown.addItem("Word", MimeTypes.docx, true);
            } else {
                fileTypeDropdown.addItem("Excel", MimeTypes.xlsx, true);
            }
            result.addWidget(fileTypeDropdown);
        }
    }

    private getFileCount() : number {
        let fileCount = 0;
        if (Context.eventData.drive.hasOwnProperty("selectedItems")) {
            fileCount = Context.eventData.drive.selectedItems.length;
        }
        return fileCount;
    }
}

class GmailAppCardFactory {
    private card : GoogleAppsScript.Card_Service.CardBuilder;
    private email : GoogleAppsScript.Gmail.GmailMessage;
    private option : string;

    constructor(option : string = null) {
        this.card = CommonCardFactory.create();
        this.option = option;
        if(Context.eventData.hasOwnProperty("gmail")) {
            this.email = GmailApp.getMessageById(Context.eventData.gmail.messageId);
        }
    }

    create() : GoogleAppsScript.Card_Service.CardBuilder {
        if(Context.eventData.hasOwnProperty("gmail")) {
            if(this.isEmailAlreadySaved()) {
                this.card.addSection(CommonCardFactory.createNotificationSection("This email has already been saved."));
            }
            this.card.addSection(this.createSection())
        } else {
            this.card.addSection(CommonCardFactory.createNotificationSection("Please open an email."));
        }
        return this.card;
    }

    private isEmailAlreadySaved() {
        let dmSystem = DMSystemRepository.getSelectedDMSystem();
        return dmSystem != null && typeof dmSystem.savedFileIds !== "undefined" && dmSystem.savedFileIds.includes(Context.eventData.gmail.messageId);
    }

    createSection() : GoogleAppsScript.Card_Service.CardSection {
        let result = CardService.newCardSection();
        if(this.emailHasAttachments()) {
            this.addWidgetsForEmailWithAttachments(result);
        }
        result.addWidget(this.createSaveButton());
        return result;
    }

    private emailHasAttachments() {
        return this.email.getAttachments().length > 0;
    }

    private addWidgetsForEmailWithAttachments(section: GoogleAppsScript.Card_Service.CardSection) {
        section.addWidget(this.createSaveOptionsDropdown());
        if (this.option == "attachments") {
            section.addWidget(this.createAttachmentCheckboxes());
        }
    }

    private createSaveOptionsDropdown() {
        return CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Options")
            .setFieldName("saveEmailOption")
            .addItem("Save email & attachments together", "merge", this.option ? this.option == "merge" : false)
            .addItem("Save attachments only", "attachments", this.option ? this.option == "attachments" : false)
            .setOnChangeAction(CardService.newAction().setFunctionName("onLoadGmail"));
    }

    private createAttachmentCheckboxes() : GoogleAppsScript.Card_Service.SelectionInput {
        let attachmentCheckboxes = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setFieldName("saveAttachments")
            .setTitle("Select attachments to save");
        for (let attachment of this.email.getAttachments()) {
            attachmentCheckboxes.addItem(attachment.getName(), attachment.getHash(), false);
        }
        return attachmentCheckboxes;
    }

    private createSaveButton(): GoogleAppsScript.Card_Service.Widget {
        let result = CardService.newTextButton();
        result.setText("Save");
        let action = CardService.newAction();
        action.setFunctionName("onClickSaveEmail");
        result.setOnClickAction(action);
        return result;
    }
}
class Context {
    public static eventData : any;
}

class FileHelper {
    public static isGoogleType(id : string) : Boolean {
        let file = DriveApp.getFileById(id);
        let fileMimeType = file.getMimeType();
        return fileMimeType === MimeTypes.gdoc || fileMimeType === MimeTypes.sheets;
    }

    public static saveToDocumentManager(fileIds : string[], dmSystem : DMSystem) : GoogleAppsScript.Card_Service.ActionResponse {
        for(let fileId of fileIds) {
            let fileData : GoogleAppsScript.Base.Blob;
            if(FileHelper.isGoogleType(fileId)) {
                let saveAsMimeType : string;
                let file = DriveApp.getFileById(fileId);
                if(Context.eventData.formInput.hasOwnProperty("fileType")) {
                    saveAsMimeType = Context.eventData.formInput.fileType;
                } else if (file.getMimeType() == MimeTypes.gdoc) {
                    saveAsMimeType = MimeTypes.docx;
                } else {
                    saveAsMimeType = MimeTypes.xlsx;
                }
                let driveExportUrl = DriveHelper.getExportUrl(fileId, saveAsMimeType);
                let driveExportOptions = DriveHelper.getExportOptions();
                fileData = UrlFetchApp.fetch(driveExportUrl, driveExportOptions).getBlob();
            } else {
                fileData = DriveApp.getFileById(fileId).getBlob();
            }
            UrlFetchApp.fetch(dmSystem.getAnonymousUploadURL(), DMSystem.getAnonymousUploadOptions(fileData));
            dmSystem.addSavedFileId(fileId);
            DMSystemRepository.update(dmSystem);

        }
        return DMSystem.createRedirectToImportPageActionResponse(dmSystem.getImportPageURL());
    }

    public static isFileEmpty(fileId : string) {
        try {
            DriveApp.getFileById(fileId);
            return false;
        } catch {
            return true;
        }
    }

}

class DriveHelper { 
    public static getExportUrl(id : string, mimeType : string) : string {
        return "https://www.googleapis.com/drive/v2/files/" + id + "/export?mimeType=" + mimeType;
    }

    public static getExportOptions() : any {
        return {
            method: "get",
            headers: {
                Authorization: "Bearer " + ScriptApp.getOAuthToken()
            }
        }; 
    }
}