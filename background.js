const browser = 'edge';  //Used to make sure that the telemetry indicates that the browser was 'edge'
chrome.action.onClicked.addListener(function (tab)  //This function fires when the 'Share to Teams' extension icon is clicked
{ 
    chrome.windows.getCurrent(function(currentWindow) 
    {
        if(tab.url == 'edge://newtab/')  
        {
            //Ignoring the click if the URL field is empty
        }
        else
        {
            let finalReferrerURL = getFinalReferrerURL(tab,'topNav')  //The referrer URL would contain 'topNav' to hint where the final Share Action happened
            var sttURL = 'https://teams.microsoft.com/share?href='+encodeURIComponent(tab.url)+'&referrer='+encodeURIComponent(finalReferrerURL); //URL ecoding both the Tab URL which is to be shared and the Refferer URL we want to create
            createWindow(sttURL, tab, currentWindow);
        }
    });
});

chrome.runtime.onInstalled.addListener(function() 
{
    chrome.contextMenus.create(
    {
        title: 'Share link to Teams',
        id: 'Link', // you'll use this in the handler function to identify this context menu item
        contexts: ['link'],
    });
    chrome.contextMenus.create(
    {
        title: 'Share to Teams',
        id: 'nonLink', // you'll use this in the handler function to identify this context menu item
        contexts: ['page'],
    });
});

chrome.contextMenus.onClicked.addListener(function(info, tab) 
{
    chrome.windows.getCurrent(function(currentWindow) 
    {
        var sttURL='';
        if (info.menuItemId === "Link") 
        {
            let finalReferrerURL = getFinalReferrerURL(tab,'linkContextMenu')
            var sttURL = 'https://teams.microsoft.com/share?href='+encodeURIComponent(info.linkUrl)+'&referrer='+encodeURIComponent(finalReferrerURL);
            createWindow(sttURL, tab, currentWindow);	
        }
        if (info.menuItemId === "nonLink") 
        { 
            if(tab.url == 'edge://newtab/')
            {
                //Ignoring the click if the URL field is empty
            }
            else
            {
                let finalReferrerURL = getFinalReferrerURL(tab,'nonLinkContextMenu')
                var sttURL = 'https://teams.microsoft.com/share?href='+encodeURIComponent(tab.url)+'&referrer='+encodeURIComponent(finalReferrerURL);
                createWindow(sttURL, tab, currentWindow);
            }
        }
    });
});

function getFinalReferrerURL(currentTab, clickLocation) 
{
	let originalTabURL = (new URL(currentTab.url));
	return originalTabURL.protocol+'//'+originalTabURL.host+'?sttsource='+browser+'-'+clickLocation;  //Creating the Referrer URL from the Tab URL, browser name and ClickLocation
}

function createWindow(sttURL, tab, currentWindow) //This function actually tries to center the 'Share to Teams' popup
{
    var w = Math.round((tab.width)/4);
    var h = Math.round((tab.height)/6);
    let createData = 
    {
        url: sttURL,
	    focused: true,
        width: w*2,
        height: h*5,
        top: currentWindow.top + h,
        left: currentWindow.left + w,
        type: 'popup'
    }
    chrome.windows.create(createData)
}