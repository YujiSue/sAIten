/*********************************************************/
/**                      Common                         **/
/*********************************************************/
///////////////////////////////////////////////////////////
/** 
 * Set Gemini API key 
 */
function setAPIKey(key=null) {
    /* Get pref. */
    const pref = getPref();
    
    /* Set key */
    if (key) props.setProperty('API_KEY', key);
    else {
        /* Display prompt to enter user's API key */
        const result = showPrompt(trans('msg.set_api_key'));
        if (result) props.setProperty('API_KEY', result);
        else return;
    }
    /* Test API */
    const response = askGemini('This is a test. Reply with only "ok"', getEndpoint(pref));

    /* Notification */
    if (response.status == STATUS_OK) {
        /* Successed */
        showToast(trans('msg.api_reg_success'));
    }
    else {
        /* Failed */
        console.error(response.message);
        props.deleteProperty("API_KEY");
        showToast(`${trans('msg.api_reg_failed')}\n${trans(response.message)}`);
    }
}
///////////////////////////////////////////////////////////
function getSafety(pref) {
    let safety = [];
    if (pref.allow_harass.value) safety.push({ "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" });
    if (pref.allow_hate.value) safety.push({ "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" });
    if (pref.allow_sexual.value) safety.push({ "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" });
    if (pref.allow_danger.value) safety.push({ "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" });
    return safety;
}
///////////////////////////////////////////////////////////
function getEndpoint(pref) {
    return props.getProperty('gemini').replace('%model%', pref.ai_model.value);
}
///////////////////////////////////////////////////////////
/** 
 * Use Gemini AI 
 */
function askGemini(prompt, endpoint, safety=[]) {
    // Get API key
    const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!apiKey) return {"status": STATUS_NG, "message" : "msg.no_api_key"};
    
    /* Gemini endpoint URL */
    const url = `${endpoint}?key=${apiKey}`;

    /* Make payload */
    let payload = {"contents": [{ "parts": [{ "text": prompt }] }]};
    if (0 < safety.length) payload["safetySettings"] = safety;

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
            console.error(`Error #${responseCode}\n${responseBody}`);
            return {"status": STATUS_ERROR, "error": responseCode, 'message':responseBody };
        }
    } catch (e) {
        console.error(`Error\n${e.toString()}`);
        return {"status": STATUS_ERROR, 'message':e.toString() };
    }
}
///////////////////////////////////////////////////////////
/*********************************************************/