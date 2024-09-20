main();

function main() {
  var folder = Folder.selectDialog("フォルダを選択してください");
  if (!folder) {
    alert("中止しました");
    return;
  }

  var dialog = new Window("dialog", "タイトルを入力してください");
  var inputField = dialog.add("edittext", undefined, "");
  inputField.characters = 30;
  var buttonGroup = dialog.add("group");
  buttonGroup.add("button", undefined, "OK", { name: "ok" });
  buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
  if (dialog.show() == 1) {
    var userInput = inputField.text === "" ? "" : inputField.text + "_";

    // エラーログファイルを作成
    var errorLogFile = function(messages) {
      var errorLog = new File(folder.fsName + "/error_log.txt");
      if (errorLog) {
        errorLog.encoding = "UTF-8";
        errorLog.open("a");
        errorLog.write(messages.join("\n") + "\n");
        errorLog.close();
      } else {
        alert("エラーログファイルが開けませんでした");
      }
    }

    // 処理中のダイアログを作成
    var dialogLoad = new Window("palette", "処理中");
    dialogLoad.add("statictext", undefined, "処理中...");
    dialogLoad.show();

    // フォルダ内の .indd ファイルとサブフォルダを処理
    processFolder(folder, userInput, errorLogFile, 1); // 深さ制限を追加

    dialogLoad.close();
    alert("処理が完了しました。\nエラーログも確認してください。");
  } else {
    alert("処理がキャンセルされました");
  }
}

function processFolder(folder, userInput, errorLogFile, depth) {
  // 深さの制限（サブフォルダのみ処理）
  if (depth > 2) {
    return;
  }

  // .indd ファイルを処理
  var files = folder.getFiles("*.indd");
  if (files.length > 0) {
    spledSplit(folder, userInput, files, errorLogFile);
    linkSyusyu(folder, errorLogFile);
  }

  // サブフォルダを処理（再帰的）
  var subFolders = folder.getFiles(function(file) {
    return file instanceof Folder;
  });

  for (var i = 0; i < subFolders.length; i++) {
    var subFolder = subFolders[i];
    
    // パッケージされたフォルダをスキップ
    if (!subFolder.name.match(/^Processed_/)) {
      processFolder(subFolder, userInput, errorLogFile, depth + 1); // 再帰的に深さを渡す
    }
  }
}

function spledSplit(folder, userInput, files, errorLogFile) {
  for (var f = 0; f < files.length; f++) {
    var doc;
    try {
      doc = app.open(files[f], false);
    } catch (e) {
      errorLogFile(["ドキュメントを開く際にエラーが発生しました: " + files[f].name + " - エラー: " + e.message]);
      continue;
    }

    if (doc.saved === false || doc.modified === true) {
      errorLogFile(["ドキュメントを保存してから実行してください: " + files[f].name]);
      doc.close(SaveOptions.NO);
      continue;
    }

    // ドキュメントページの移動を許可しない
    doc.documentPreferences.allowPageShuffle = false;
    var spr_obj = doc.spreads;

    for (var i = 0, iL = spr_obj.length; i < iL; i++) {
      spr_obj[i].allowPageShuffle = false;
      var pag_obj = spr_obj[i].pages;
      var spr_name = pag_obj.length > 1 ? "0" + pag_obj[0].name + "-0" + pag_obj[pag_obj.length - 1].name : pag_obj[0].name;
      spr_obj[i].insertLabel('sp_id', "" + i);
      spr_obj[i].insertLabel('sp_name', spr_name);
      spr_obj[i].insertLabel('start_p_num', pag_obj[0].name.replace(pag_obj[0].appliedSection.name, ""));
    }

    var org_doc_path = doc.fullName;
    var iiL = spr_obj.length;
    var new_fd_path = folder.fsName + "/00_SplitData/";
    Folder(new_fd_path).create();

    // バックアップを作成
    doc.close(SaveOptions.NO, new File(org_doc_path + ".bk"));

    // 再度バックアップを開く
    for (var is = 0; is < iiL; is++) {
      var doc2 = app.open(File(org_doc_path + ".bk"), false);
      var spr_obj2 = doc2.spreads;

      // 他のスプレッドを削除
      for (var ii = iiL - 1; ii >= 0; ii--) {
        if (spr_obj2[ii].extractLabel('sp_id') !== "" + is) {
          spr_obj2[ii].remove();
        }
      }

      // ページ番号のスタートを設定
      doc2.sections[0].continueNumbering = false;
      doc2.sections[0].pageNumberStart = doc2.spreads[0].extractLabel('start_p_num') * 1;

      // パッケージ機能を使用して保存
      var packageFolder = new Folder(new_fd_path + userInput + spr_obj2[0].extractLabel('sp_name'));
      packageFolder.create();
      doc2.packageForPrint(packageFolder, true, true, false, false, false, false, false);

      // 保存先,Fonts,Links,指示書,更新,非表示レイヤー,プリフライトエラー,レポート
      doc2.close(SaveOptions.NO, new File(packageFolder.fsName + "/" + userInput + spr_obj2[0].extractLabel('sp_name') + ".indd"));
    }

    // .bk ファイルを削除
    File(org_doc_path + ".bk").remove();

    function removeBkFiles(folder) {
      var files = folder.getFiles();
      for (var i = 0; i < files.length; i++) {
        if (files[i] instanceof Folder) {
          removeBkFiles(files[i]);
        } else if (files[i].name.match(/\.bk$/)) {
          files[i].remove();
        }
      }
    }

    var splitDataFolder = new Folder(folder.fsName + "/00_SplitData/");
    if (splitDataFolder.exists) {
      removeBkFiles(splitDataFolder);
    }
  }
}

function linkSyusyu(folder, errorLogFile) {
  processFolders(new Folder(folder.fsName + "/00_SplitData/"), errorLogFile);
}

function processFolders(folder, errorLogFile) {
  var subFolders = folder.getFiles(function(file) {
    return file instanceof Folder;
  });

  for (var i = 0; i < subFolders.length; i++) {
    var subFolder = subFolders[i];
    
    var inddFiles = subFolder.getFiles("*.indd");
    var linksFolder = new Folder(subFolder.fsName + "/Links");
    if (!linksFolder.exists) {
      linksFolder.create();
    }

    for (var j = 0; j < inddFiles.length; j++) {
      processDocument(inddFiles[j], linksFolder, errorLogFile);
    }
  }
}

function processDocument(docFile, linksFolder, errorLogFile) {
  var doc;
  try {
    doc = app.open(docFile, false);
  } catch (e) {
    errorLogFile(["ドキュメントのオープン中にエラーが発生しました: " + docFile.name + " - エラー: " + e.message]);
    return;
  }
  collectLinks(doc, linksFolder, errorLogFile);
  doc.close(SaveOptions.NO);
}

function collectLinks(doc, linksFolder, errorLogFile) {
  var links = doc.links;
  if (!links || links.length === 0) {
    return;
  }
  var collectedLinks = {};
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    var sourceFile = new File(link.filePath);
    if (sourceFile.exists) {
      collectSubLinksFromFile(sourceFile, linksFolder, collectedLinks, errorLogFile, doc.name);
    } else {
      errorLogFile(["リンクが見つかりません: " + link.filePath + " (ドキュメント: " + doc.name + ")"]);
    }
  }
}

function collectSubLinksFromFile(file, linksFolder, collectedLinks, errorLogFile, docTitle) {
  var subLinks = getLinks(file.fsName, errorLogFile);
  var errorMessages = [];

  for (var j = 0; j < subLinks.length; j++) {
    var subLinkName = subLinks[j];
    if (!collectedLinks[subLinkName]) {
      var subLinkFile = new File(file.path + "/" + subLinkName);
      if (subLinkFile.exists) {
        var subDestFile = new File(linksFolder.fsName + "/" + subLinkName);
        try {
          copyFile(subLinkFile.fsName, subDestFile.fsName);
          collectedLinks[subLinkName] = true;
          collectSubLinksFromFile(subLinkFile, linksFolder, collectedLinks, errorLogFile, docTitle);
        } catch (e) {
          errorMessages.push("ファイルのコピー中にエラーが発生しました: " + e.message);
        }
      } else {
        errorMessages.push("サブリンクが見つかりません: " + decodeURI(subLinkFile.fsName));
      }
    }
  }

  if (errorMessages.length > 0) {
    errorLogFile(["ドキュメント: " + docTitle + "\n" + errorMessages.join("\n")]);
  }
}

//リンクを移動させる
function copyFile(moto, saki) {
  var destFile = new File(saki);
  if (destFile.exists) {
    return;
  }
  var d1 = app.documents.add();
  try {
    var p = d1.spreads[0].place(File(moto));
    p[0].itemLink.copyLink(destFile);
  } catch (e) {
    alert("ファイルのコピー中にエラーが発生しました: " + e.message);
  } finally {
    d1.close(SaveOptions.NO);
  }
}

// function getLinks(filePath, errorLogFile) {
//   try {
//     var prop = "Manifest";
//     var ns = "http://ns.adobe.com/xap/1.0/mm/";
//     var xmpFile = new XMPFile(filePath, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
//     var xmpPackets = xmpFile.getXMP();
//     var xmp = new XMPMeta(xmpPackets.serialize());
//     var result = [];
//     for (var i = 1; i <= xmp.countArrayItems(ns, prop); i++) {
//       var str = xmp.getProperty(ns, prop + "[" + i + "]/stMfs:reference/stRef:filePath").toString();
//       result.push(str.slice(str.lastIndexOf("/") + 1));
//     }
//     return result;
//   } catch (e) {
//     errorLogFile(["XMP処理中にエラーが発生しました: " + filePath + " - エラー: " + e.message]);
//     return [];
//   }
// }

function getLinks(filePath, errorLogFile) {
  try {

  var prop = "Manifest";
  var ns = "http://ns.adobe.com/xap/1.0/mm/";
  if (typeof xmpLib === 'undefined') {
      var xmpLib = new ExternalObject('lib:AdobeXMPScript');
  }
  var xmpFile = new XMPFile(filePath, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
  var xmpPackets = xmpFile.getXMP();
  var xmp = new XMPMeta(xmpPackets.serialize());
  var str = "";
  var result = [];
  for (var i = 1; i <= xmp.countArrayItems(ns, prop); i++) {
      str = xmp.getProperty(ns, prop + "[" + i + "]/stMfs:reference/stRef:filePath").toString();
      result.push(str.slice(str.lastIndexOf("/") + 1));
  }
  return result;

  } catch (e) {
    errorLogFile(["XMP処理中にエラーが発生しました: " + filePath + " - エラー: " + e.message]);
    return [];
  }
}
function daialogTojiru() {
  alert("処理が完了しました");
}
