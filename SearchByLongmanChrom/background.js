// chrome.contextMenus.create({
//     "title": "search '%s' on google dic",
//     "type": "normal",
//     "contexts": ["selection"],
//     "onclick": function(info) {
//         chrome.tabs.create({
//             //url: 'https://www.ldoceonline.com/dictionary/'+encodeURIComponent(info.selectionText)
//             url: 'https://translate.google.co.jp/?hl=ja&sl=en&tl=ja&text=' +encodeURIComponent(info.selectionText)
//         });
//     }
// });


const copyToClipboard = (tab, text) => {
  function injectedFunction(text) {
    try {
      navigator.clipboard.writeText(text);
      //console.log('successfully');
    } catch (e) {
      //console.log('failed');
    }
  }
  chrome.scripting.executeScript({
    target: {tabId: tab.id},
    func: injectedFunction,
    args: [text]
  });
};

const updateContextMenus = async () => {
  await chrome.contextMenus.removeAll();

  chrome.contextMenus.create({
      id: "context-copytab-title-url",
      title: "search en2jp",
      contexts: ["all"]
  });
//   chrome.contextMenus.create({
//       id: "context-copytab-title",
//       title: "タブのタイトルをコピー",
//       contexts: ["all"]
//   });
//   chrome.contextMenus.create({
//       id: "context-copytab-url",
//       title: "タブのURLをコピー",
//       contexts: ["all"]
//   });
};


const searchWord = info =>
  chrome.tabs.create({
        url: 'https://translate.google.co.jp/?hl=ja&sl=en&tl=ja&text=' +encodeURIComponent(info.selectionText)
  });

chrome.runtime.onInstalled.addListener(updateContextMenus);
chrome.runtime.onStartup.addListener(updateContextMenus);
chrome.contextMenus.onClicked.addListener((info, tab) => {
  switch (info.menuItemId) {
  case 'context-copytab-title-url':
    searchWord(info);
    //copyToClipboard(tab, tab.title+'\n'+tab.url);
    break;
  case 'context-copytab-title':
    copyToClipboard(tab, tab.title);
    break;
  case 'context-copytab-url':
    copyToClipboard(tab, tab.url);
    break;
  }
});


