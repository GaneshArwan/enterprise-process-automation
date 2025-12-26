/**
 * Factory class for generating script generators based on script types.
 */
class ScriptFactory {
    static getGenerator(scriptType) {
        const generators = {
            [ScriptTypes.SCRIPT_TOYS_CONS]: ScriptGenerators.generateToysCons,
            [ScriptTypes.NO_SITE]: ScriptGenerators.generateNoSite,
            [ScriptTypes.UPDATE_PIR_AND_CURRENCY]: ScriptGenerators.generateUpdatePIRAndCurrency,
            [ScriptTypes.UPDATE_PIR_NEW]: ScriptGenerators.generateUpdatePIRNew,
            [ScriptTypes.CURRENCY_NEW]: ScriptGenerators.generateCurrencyNew,
            [ScriptTypes.SCRIPT_HOME_ESSENTIALS]: ScriptGenerators.generateScriptHomeEssentials,
            [ScriptTypes.SCRIPT_RETAIL_NEW]: ScriptGenerators.generateScriptRetailNew,
            [ScriptTypes.SCRIPT_SAP_A]: ScriptGenerators.generateScriptSapA,
            [ScriptTypes.SCRIPT_SAP_B]: ScriptGenerators.generateScriptSapB,
            [ScriptTypes.SCRIPT_MODIFY_PRICE]: ScriptGenerators.generateScriptModifyPrice,
            [ScriptTypes.CREATE_PIR_SITE]: ScriptGenerators.generateCreatePIRSite,
        };

        return generators[scriptType];
    }
}

/**
 * Utility class for generating scripts based on row values.
 */
class ScriptGenerators {
    static processRowValues(row) {
        return Object.fromEntries(
            Object.entries(row).map(([key, val]) => {
                const newKey = key === 'PURCHASE_ORGANIZATION' ? 'PURCHASING_ORGANIZATION' : key;
                return [newKey, (val ?? '').toString().trim()];
            })
        );
    }

    static generateScriptHomeEssentials(row) {
        const processedRow = ScriptGenerators.processRowValues(row);
        // Standard SAP VBScript generation
        return 'session.findById("wnd[0]").maximize\n' +
            'session.findById("wnd[0]/tbar[0]/okcd").text = "ME11"\n' +
            'session.findById("wnd[0]").sendVKey(0)\n' +
            'session.findById("wnd[0]/usr/ctxt[0]").text = "' + processedRow["VENDOR_CODE"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[1]").text = "' + processedRow["ARTICLE_NUMBER"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[2]").text = "' + processedRow["PURCHASING_ORGANIZATION"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[4]").setFocus\n' +
            'session.findById("wnd[0]").sendVKey(0)\n' +
            'session.findById("wnd[0]/usr/txt[5]").text = "' + processedRow["PLAN_DELIVERY_TIME_(DAYS)"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[5]").text = "' + processedRow["PURCHASING_GROUP"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[8]").text = "' + processedRow["STANDARD_PO_QTY"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[7]").text = "' + processedRow["CONFIRMATION_CONTROL"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[9]").text = "' + processedRow["MINIMUM_ORDER_QTY"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[9]").text = "' + processedRow["TAX_CODE"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[15]").text = "' + processedRow["NET_PRICE"] + '"\n' +
            'session.findById("wnd[0]/tbar[0]/btn[11]").press\n' +
            'session.findById("wnd[0]/tbar[0]/btn[3]").press\n';
    }

    static generateScriptSapA(row) {
        const processedRow = ScriptGenerators.processRowValues(row);
        return 'session.findById("wnd[0]").maximize\n' +
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11"\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = "' + processedRow.VENDOR_CODE + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = "' + processedRow.ARTICLE_NUMBER + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = "' + processedRow.PURCHASING_ORGANIZATION + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINE-EKORG").setFocus\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/chkEINA-RELIF").selected = true\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/txtEINE-APLFZ").text = "' + processedRow["PLAN_DELIVERY_TIME_(DAYS)"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINE-EKGRP").text = "' + processedRow.PURCHASING_GROUP + '"\n' +
            'session.findById("wnd[0]/usr/txtEINE-NORBM").text = "' + processedRow.STANDARD_PO_QTY + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINE-BSTAE").text = "' + processedRow.CONFIRMATION_CONTROL + '"\n' +
            'session.findById("wnd[0]/usr/txtEINE-MINBM").text = "' + processedRow.MINIMUM_ORDER_QTY + '"\n' +
            'session.findById("wnd[0]/usr/ctxtEINE-MWSKZ").text = "' + processedRow.TAX_CODE + '"\n' +
            'session.findById("wnd[0]/usr/txtEINE-NETPR").text = "' + processedRow.NET_PRICE + '"\n' +
            'session.findById("wnd[0]").sendVKey 11\n';
    }

    static generateRetailConsScript(row) {
        const processedRow = ScriptGenerators.processRowValues(row);
        return 'session.findById("wnd[0]").maximize\n' +
            'session.findById("wnd[0]/tbar[0]/okcd").text = "me11"\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/ctxt[0]").text = "' + processedRow["VENDOR_CODE"] + '"\n' + 
            'session.findById("wnd[0]/usr/ctxt[1]").text = "' + processedRow["ARTICLE_NUMBER"] + '"\n' + 
            'session.findById("wnd[0]/usr/ctxt[2]").text = "' + processedRow["PURCHASING_ORGANIZATION"] + '"\n' +
            'session.findById("wnd[0]/usr/rad[3]").setFocus\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/txt[5]").text = "' + processedRow["CONFIRMATION_CONTROL"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[5]").text = "' + processedRow["PURCHASING_GROUP"] + '"\n' + 
            'session.findById("wnd[0]/usr/txt[8]").text = "' + processedRow["STANDARD_PO_QTY"] + '"\n' + 
            'session.findById("wnd[0]/usr/ctxt[7]").text = "' + processedRow["NET_PRICE"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[9]").text = "' + processedRow["MINIMUM_ORDER_QTY"] + '"\n' + 
            'session.findById("wnd[0]/usr/ctxt[9]").text = "' + processedRow["TAX_CODE"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[15]").text = "1"\n' +
            'session.findById("wnd[0]/tbar[1]/btn[8]").press\n' +
            'session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txt[2,0]").text = "0"\n' +
            'session.findById("wnd[0]/tbar[0]/btn[3]").press\n' +
            'session.findById("wnd[0]/tbar[0]/btn[11]").press\n' +
            'session.findById("wnd[0]/tbar[0]/btn[3]").press\n';
    }

    static generateNoSite(row) {
        const processedRow = ScriptGenerators.processRowValues(row);
        return 'session.findById("wnd[0]").maximize\n' +
            'session.findById("wnd[0]/tbar[0]/okcd").text = "me11"\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/ctxt[0]").text = "' + processedRow["VENDOR_CODE"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[1]").text = "' + processedRow["ARTICLE_NUMBER"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[2]").text = "' + processedRow["PURCHASING_ORGANIZATION"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[2]").setFocus\n' +
            'session.findById("wnd[0]/usr/ctxt[2]").caretPosition = 4\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]").sendVKey 0\n' +
            'session.findById("wnd[0]/usr/txt[5]").text = "' + processedRow["PLAN_DELIVERY_TIME_(DAYS)"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[5]").text = "' + processedRow["PURCHASING_GROUP"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[8]").text = "' + processedRow["STANDARD_PO_QTY"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[7]").text = "' + processedRow["CONFIRMATION_CONTROL"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[9]").text = "' + processedRow["MINIMUM_ORDER_QTY"] + '"\n' +
            'session.findById("wnd[0]/usr/ctxt[9]").text = "' + processedRow["TAX_CODE"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[15]").text = "' + processedRow["NET_PRICE"] + '"\n' +
            'session.findById("wnd[0]/usr/txt[15]").setFocus\n' +
            'session.findById("wnd[0]/usr/txt[15]").caretPosition = 14\n' +
            'session.findById("wnd[0]/tbar[0]/btn[11]").press\n' +
            'session.findById("wnd[0]/tbar[0]/btn[3]").press\n';
    }

    // ... (Keep generic generators like UpdatePIRNew, CurrencyNew as they are safe)
    
    constructor(Activity) {
        // ... (rest of the class remains same)
    }
}

/**
 * Class for generating PIR (Purchase Invoice Request) scripts.
 */
class ScriptGenerator {
    constructor(Activity) {
        this.activity = Activity;
    }

    _getScriptName() {
        const { SCRIPT_TYPE } = this.activity.getActivityValueMap();
        return `Script ${SCRIPT_TYPE} ${new Date().toISOString()}.vbs`;
    }

    createScriptDoc(scriptString) {
        const companyName = this.activity.getCompanyName();
        const folderUID = getRequestConfig(companyName)['EXTEND_PIR_SCRIPT'];
        const folder = DriveApp.getFolderById(folderUID);
        const name = this._getScriptName();

        const docURL = createTxtFile(name, scriptString, folder);
        return docURL.getUrl();
    }

    initialSAPscript() {
        return "If Not IsObject(application) Then\n" +
            "   Set SapGuiAuto  = GetObject(\"SAPGUI\")\n" +
            "   Set application = SapGuiAuto.GetScriptingEngine\n" +
            "End If\n" +
            "If Not IsObject(connection) Then\n" +
            "   Set connection = application.Children(0)\n" +
            "End If\n" +
            "If Not IsObject(session) Then\n" +
            "   Set session    = connection.Children(0)\n" +
            "End If\n" +
            "If IsObject(WScript) Then\n" +
            "   WScript.ConnectObject session,     \"on\"\n" +
            "   WScript.ConnectObject application, \"on\"\n" +
            "End If\n"
    }

    generateScript(pirValues) {
        if (!Array.isArray(pirValues) || pirValues.length === 0) return null;

        const { SCRIPT_TYPE } = this.activity.getActivityValueMap();
        const generatorFunction = ScriptFactory.getGenerator(SCRIPT_TYPE);

        if (typeof generatorFunction !== "function") {
            console.warn(`No generator function found for script type: ${SCRIPT_TYPE}`);
            return null;
        }
        const sapScriptHeader = this.initialSAPscript();
        const scriptBody = pirValues
            .filter(rowMap => rowMap && Object.values(rowMap).some(value => typeof value === "string" && value.trim() !== ""))
            .map(rowMap => generatorFunction(rowMap))
            .join('');

        if (!scriptBody.trim()) return null;
        const scriptString = sapScriptHeader + scriptBody;
        const scriptDocURL = this.createScriptDoc(scriptString);

        return scriptDocURL;
    }
}