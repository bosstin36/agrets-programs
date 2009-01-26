// Pixiv Translator
// Created by Agret
// 
// Source code completey ripped from:
//
// Kaskus De-Obfuscator
// Created by Pandu E Poluan {http://userscripts.org/users/71414/}
// Version:      0.4.1
//
// ==UserScript==
// @name          Pixiv Translator
// @namespace     about:robots
// @description   Translates Pixiv (mostly) into English.
// @include       *pixiv.net*
// ==/UserScript==

(function () {

var replacements, regex, key, thenodes, node, s;

replacements = {
  "新規登録": "Register",
  "パスワード": "Password",
  "次回から自動的にログイン": "Remember my password",
  "※pixiv ID・Passwordを忘れた": "Forogot Pixiv ID / Password",
  "IDまたはPasswordを忘れてしまった": "Forgot Pixiv ID / Password",
  "ホーム": "Home",
  "pixivについて": "About",
  "ヘルプ": "Help",
  "開発者ブログ": "Blog",
  "お問い合わせ": "Support",
  "ID、パスワードが正しいかチェックしてください。1": "Invalid Password",
  "メールアドレス": "E-Mail",
  "広告掲載": "Contact",
  "ガイドライン": "Guidelines",
  "プライバシーポリシー": "Privacy Policy",
  "利用規約": "Terms of Use",
  "運営会社": "Operator",
  "人材募集": "Jobs",
  "お知らせ": "News",
  "このページの上部へ": "Back to Top",
  "忘れてしまったものを選択し、登録E-Mailを入力して送信してください。": "Select what you have forgotten and then enter your",
  "E-Mail宛に": "e-mail.",
  "pixiv IDの場合、pixiv IDが記載されたメールが送信されます。": "Your Pixiv ID will be sent to your email address.",
  "Passwordの場合、Password再設定ページのURLが記載されたメールを送信します。": "If you have forgotten your password instructions will be sent on how to reset it.",
  "pixiv IDを忘れた": "Pivix ID",
  "Passwordを忘れた": "Password",
  };
regex = {};
for ( key in replacements ) {
  regex[key] = new RegExp(key, 'gi');
}

// Now, retrieve the text nodes
thenodes = document.evaluate( "//text()" , document , null , XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE , null );

// Perform a replacement over all the nodes
for ( var i=0 ; i<thenodes.snapshotLength ; i++ ) {
  node = thenodes.snapshotItem(i);
  s = node.data;
  for ( key in replacements ) {
    s = s.replace( regex[key] , replacements[key] );
  }
  node.data = s;
}

// Now, retrieve the A nodes
thenodes = document.evaluate( "//a" , document , null , XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE , null );

// Finally, perform a replacement over all A nodes
for ( var i=0 ; i<thenodes.snapshotLength ; i++ ) {
  node = thenodes.snapshotItem(i);
  // Here's the key! We must replace the "href" instead of the "data"
  s = node.href;
  for ( key in replacements ) {
    s = s.replace( regex[key] , replacements[key] );
  }
  node.href = s;
}

// Now, retrieve the input nodes
thenodes = document.evaluate( "//input" , document , null , XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE , null );

// Finally, perform a replacement over all A nodes
for ( var i=0 ; i<thenodes.snapshotLength ; i++ ) {
  node = thenodes.snapshotItem(i);
  // Here's the key! We must replace the "href" instead of the "data"
  s = node.value;
  for ( key in replacements ) {
    s = s.replace( regex[key] , replacements[key] );
  }
  node.value = s;
}

})();
