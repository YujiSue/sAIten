/*********************************************************/
/**                      Common                         **/
/*********************************************************/

///////////////////////////////////////////////////////////
/** Set Gemini API key */
function setAPIKey(key=null) {
    /* Get general info. */
    var props = PropertiesService.getScriptProperties();
    const info = getInfo();
    const pref = getPreference();
    const loc = new Localizer();
    /* Set key */
    if (key) props.setProperty("API_KEY", key);
    else {
        /* Display prompt to enter user's API key */
        const ui = SpreadsheetApp.getUi();
        const result = ui.prompt(info.set_api_msg, ui.ButtonSet.OK_CANCEL);
        if (result.getSelectedButton() === ui.Button.OK) 
          props.setProperty("API_KEY", result.getResponseText().trim());
    }
    /* Test API */
    const res = askGemini('This is a test. Reply with only "ok"', 
                            info.api_ep.replace('%model%', pref.ai_model));
    if (res.status == STATUS_OK) {
        /* Successed */
        showToast(loc.trans('api_registration_success'));
    }
    else {
        /* Failed */
        props.deleteProperty("API_KEY");
        showToast(loc.trans('api_registration_fail'));
    }
}
///////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////
/** Use Gemini AI */
function askGemini(prompt, endpoint, safety=[]) {
    // Get API key
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return {"status": STATUS_NG, "msg" : "no_api_key_error"};
    
    /* Gemini endpoint URL */
    const url = `${endpoint}?key=${apiKey}`;

    /* Make payload */
    const payload = 0 < safety.length ? 
                    {"contents": [{ "parts": [{ "text": prompt }] }], "safetySettings": safety} :
                    {"contents": [{ "parts": [{ "text": prompt }] }]};
    //
    const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': false
    };
    /* Call API */
    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        //
        if (responseCode === 200) {
            const jsonResponse = JSON.parse(responseBody);
            // Extract AI's answer
            if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
                jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
                jsonResponse.candidates[0].content.parts.length > 0) {
                return {"status":STATUS_OK, "text":jsonResponse.candidates[0].content.parts[0].text.trim()};
            } 
            // 
            else return { "status": STATUS_NG };
        } else {
            return {"status": STATUS_ERROR, "error": responseCode, 'message':responseBody };
        }
    } catch (e) {
        return {"status": STATUS_ERROR, 'msg':e.toString() };
    }
}
///////////////////////////////////////////////////////////
/*********************************************************/