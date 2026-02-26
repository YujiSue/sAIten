

/** 
 * Use Gemini AI 
 */
function askGemini(prompt, pref, prop) {
    // Get API key
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return {"status": STATUS_ERROR, "error" : "API key is not set. Please set the key first."};
    
    // Gemini endpoint URL
    const url = `${pref.aimodel.replace('<model>', pref.aimodel)}?key=${apiKey}`;
    //
    let safetyOp = [];
    if (props.allow_harass) safetyOp.push({ "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" });
    if (props.allow_hate) safetyOp.push({ "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" });
    if (props.allow_sexual) safetyOp.push({ "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" });
    if (props.allow_danger) safetyOp.push({ "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" });
    // Make payload
    const payload = 0 < safetyOp.length ? 
                    {"contents": [{ "parts": [{ "text": prompt }] }], "safetySettings": safetyOp} :
                    {"contents": [{ "parts": [{ "text": prompt }] }]};
    //
    const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': false
    };
    /** Call API */
    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        if (responseCode === 200) {
            const jsonResponse = JSON.parse(responseBody);
            // Extract AI's answer
            if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
                jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
                jsonResponse.candidates[0].content.parts.length > 0) {
                return {"status":STATUS_OK, "text":jsonResponse.candidates[0].content.parts[0].text.trim()};
            } 
            // 
            else return {"status": STATUS_ERROR, "error": locale.msg.api_fail[props.langs] };
                //return "エラー: Geminiからの有効な回答がありません。不適切な回答を含んでいた可能性があります。";
        } else {
            return {"status": STATUS_ERROR, "error": `${locale.msg.api_call_error[props.lang]} (# ${responseCode})\n${responseBody}`} ;
        }
    } catch (e) {
        return `${locale.msg.api_call_error[props.lang]}\n${e.toString()}`;
    }
}
