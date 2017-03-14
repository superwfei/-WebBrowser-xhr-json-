using mshtml;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;

namespace TYC_Gather
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [ComVisible(true)]
    public partial class TYC_V2 : Form
    {
        private string result = string.Empty;
        private string script = @"
;(function () {
    if (typeof window.CustomEvent === 'function') return false;
    function CustomEvent ( event, params ) {
        params = params || { bubbles: false, cancelable: false, detail: undefined };
        var evt = document.createEvent('CustomEvent');
        evt.initCustomEvent( event, params.bubbles, params.cancelable, params.detail );
        return evt;
    }
    CustomEvent.prototype = window.Event.prototype;
    window.CustomEvent = CustomEvent;
    window._xhrs = window._xhrs||{count:0,result:[],send:0,xhr:0,oldXHR:window.XMLHttpRequest};
})();
;(function()
{
    function ajaxEventTrigger(event) {
        var ajaxEvent = new CustomEvent(event, { detail: this });
        window.dispatchEvent(ajaxEvent);
        if (event=='ajaxLoadEnd'){
            if (/^<[\s\S]*>\s*$/gi.test(this['responseText'])){
                //alert(this['responseText'])
                window._xhrs['result'].push('{""url"":""' + this['Url'] + '"",""result"":null}');
            }else{
                window._xhrs['result'].push('{""url"":""' + this['Url'] + '"",""result"":' + (this['responseText'] || 'null') + '}');
                window.external.sendToWinForm(this['Url'],this['responseText'],event);
            }
            returnCSharp();
        }
    }
    var oldXHR = window._xhrs['oldXHR'] || window.XMLHttpRequest;
    function newXHR()
    {
        var realXHR = new oldXHR();
        window._xhrs['xhr']++;
        realXHR.addEventListener('abort', function() { ajaxEventTrigger.call(this, 'ajaxAbort'); }, false);
        realXHR.addEventListener('error', function() { ajaxEventTrigger.call(this, 'ajaxError'); }, false);
        realXHR.addEventListener('load', function() { ajaxEventTrigger.call(this, 'ajaxLoad'); }, false);
        realXHR.addEventListener('loadstart', function() { ajaxEventTrigger.call(this, 'ajaxLoadStart'); }, false);
        realXHR.addEventListener('progress', function() { ajaxEventTrigger.call(this, 'ajaxProgress'); }, false);
        realXHR.addEventListener('timeout', function() { ajaxEventTrigger.call(this, 'ajaxTimeout'); }, false);
        realXHR.addEventListener('loadend', function() { ajaxEventTrigger.call(this, 'ajaxLoadEnd'); }, false);
        realXHR.addEventListener('readystatechange', function() { ajaxEventTrigger.call(this, 'ajaxReadyStateChange'); }, false);
        var oldOpen = realXHR.open;
        realXHR.open = function(method,url){
            window._xhrs['count']++;
            this.Url = url;
            return oldOpen.apply(this,arguments);
        }
        return realXHR;
    }
    window.XMLHttpRequest = newXHR;
})();
function returnCSharp(){
    var xel = document.getElementById('live-search');
    if (xel){
        xel.value = window._xhrs.count + ' : ' + window._xhrs.result.length + ' : ' + window._xhrs['xhr'];
    }
    if (window._xhrs.count == window._xhrs.result.length && !window._xhrs['send']){
        window._xhrs['send']=1;
        setTimeout(function(){
            if (window._xhrs.count == window._xhrs.result.length){
                window.external.sendToCSharp(window._xhrs['result']);
            }else{
                window._xhrs['send']=0;
                returnCSharp();
            }
        },500);
    }
}
            ";
        private string _id = string.Empty;
        private SHDocVw.WebBrowser wb = null;
        public string Result
        {
            get
            {
                return result;
            }
        }
        public string TycID
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
                webBrowser1.Navigate("http://www.tianyancha.com/company/" + _id);
            }
        }
        public TYC_V2()
        {
            InitializeComponent();
            webBrowser1.ObjectForScripting = this;
            webBrowser1.ScriptErrorsSuppressed = true;
            wb = (SHDocVw.WebBrowser)webBrowser1.ActiveXInstance;
            wb.BeforeScriptExecute += Wb_BeforeScriptExecute;
            webBrowser1.DocumentCompleted += WebBrowser1_DocumentCompleted;
        }
        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string html = webBrowser1.DocumentText;
            if (html.IndexOf("为确认本次访问为正常用户行为，请您协助验证。") > 0)
            {
                this.Close();
                return;
            }
        }
        private void Wb_BeforeScriptExecute(object pDispWindow)
        {
            try
            {
                IHTMLDocument2 vDocument = (IHTMLDocument2)webBrowser1.Document.DomDocument;
                vDocument.parentWindow.execScript(script, "javascript");
            }
            catch (Exception ex)
            {
                this.Close();
            }
        }
        public void sendToWinForm(string url, string body, string stat)
        {
            dataGridView1.Rows.Add(new object[] { url, body, stat });
        }
        public void sendToCSharp(string obj)
        {
            result = obj;
            this.Close();
        }
    }
}
