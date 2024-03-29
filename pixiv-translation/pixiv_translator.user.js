﻿// Pixiv Translator
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
  "新規":"New",
  "登録":"Register",
  "パスワード":"Password",
  "の再設定":"Reset",
  "pixiv IDと新しいパスワードを入力して送信してください。":"Enter Pixiv ID and new password to continue",
  "確認":"Check", // Reset password
  "について":"About",
  "ページの":"Page",
  "画像の":"Image",
  "次回から自動的にログイン": "Remember my password",
  "※":"*",
  "pixiv ID・Passwordを忘れた": "Forogot Pixiv ID / Password",
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
  "上部へ":"Top",
  "この":"This",
  "忘れてしまったものを選択し、登録E-Mailを入力して送信してください。": "Select what you have forgotten and then enter your",
  "E-Mail宛に": "e-mail.",
  "pixiv IDの場合、pixiv IDが記載されたメールが送信されます。": "Your Pixiv ID will be sent to your email address.",
  "Passwordの場合、Password再設定ページのURLが記載されたメールを送信します。": "If you have forgotten your password instructions will be sent on how to reset it.",
  "pixiv IDを忘れた": "Pivix ID",
  "Passwordを忘れた": "Password",
  "送　信": "Submit",
  "ログイン": "Login",
  "入力していただいたE-Mailにメールを送信しました。": "Please check your e-mail.",
  "戻る": "Back",
  "必要な情報が見つかりませんでした": "Unable to find required information.",
  "プロフィール":"Profile",
  "トップ":"Profile",
  "作業環境":"Work Environment", // ????
  "企画目録":"Inventory Planning", // ????
  "変更":"Change",
  "プロフィール確認":"My Profile",
  "プロフィールを見る":"View Profile",
  "掲示板を見る":"Users Board",
  "pixivに招待する":"Invite a Friend",
  "友人をInvite a Friend":"Invite a Friend",
  "新たにpixivに招待したい友人・知人のE-Mailを記入してください。":"Invite a friend by e-mail",
  "相手のE-Mail":"Friends e-mail",
  "あなたのお名前":"Your Name",
  "メッセージ":"Message",
  "を送る":"Send",
  "宛　先":"Recipient",
  "件　名":"Subject",
  "確認画面":"Send",
  "みんなの新着":"Newest",
  "イラスト":"Image",
  "ーランキング":"Ranking",
  "デイリ ":"Daily",
  "ウィークリ ":"Weekly",
  "投稿日時":"Post Date",
  "現在の速度":"Current Speed",
  "パーソナル":"Personal",
  "使用ツール":"Tools",
  "人気のタグ":"Tags",
  "タグ":" Tags",
  "イラストの投稿":"Upload",
  "設定変更":"Settings",
  "ログアウト":"Logout",
  "全選択":"Select All",
  "選択解除":"Deselect All",
  "投稿者":"From",
  "あなたの":"Your",
  "解除":"Delete",
  "公開":"Visible ",
  "非":"Non-",
  "設定":"Setting",
  "非公開にする":"Make Private",
  "公開にする":"Make Public",
  "スペース区切りで複数入力できます。英数字等は半角に統一されます。":"You may add multiple space seperated tags",
  "詳細画面に反映されるまで時間がかかることがあります。":"It may take some time to become visible",
  "ブックマークタグ":"Add Tags",
  "に追加":"Add to",
  "閲覧数":"Views",
  "回数":"Times",
  "総合点":"Total Rating",
  "お気に入り":"Favorites", // For "Add to favorites" on an image
  "ブックマーク":"Favorites",
  "を見る":"User",
  "文字以内":"Character Limit",
  "関連":"Related",
  "編集":"Add",
  "追加する":"Add",
  "コメント":"Comments",
  "評価":"Rating",
  "履歴":" History",
  "もっと見る":"More",
  "はありません":"No",
  "タイトル・キャプション":"Title",
  "検　索":"Search",
  "検索":"Search",
  "結果":"Results",
  "投稿順":"By Date",
  "ランダム":"Random",
  "で選んだ":"choice of",
  "管理":"Manage",
  "イラストの":"Pictures",
  "再検索":"Search",
  "ピックアップ":"Selection",
  "お気に入りユーザー":"Watched By",
  "あなたをお気に入りに登録しているユーザーはいません":"There is nobody watching you.",
  "あなたをお気に入りに登録しているユーザー":"Users Watching You",
  "あなたを登録しているユーザー":"Your Watched Users",
  "ユーザ":"User",
  "非公開中":"Private",
  "公開中":"Public",
  // May as well translate some image tags
  "オリジナル":"Original",
  "キャラクター":"Character",
  "羽":"Feather",
  "翼":"Wing",
  "セカンドライフ":"Second Life",
  "獣":"Animal", // 獣人 for "Animal People", furries!
  "人":"People",
  "竜":"Dragon", // 竜人 for "Dragon People"
  "龍":"Dragon",
  "コードギアス":"Code Geass",
  "ルルC":"Lulu C",
  "狼":"Wolf",
  "青":"Green",
  "鎧":"Armor",
  "甲胄":"Armor",
  "女の子":"Girl",
  "ウィザード":"Wizard",
  "精灵":"Wizard",
  "オンラインゲーム":"Online Games",
  "ナルト":"Naruto",
  "ルルーシュ":"Lelouch",
  "交流絵":"Picture Exchange",
  "ヌード":"Nude",
  "おっぱい":"Boobs",
  "尻":"Ass",
  "神様":"God",
  "版権":"Copyright",
  // Fix spacing and ordering- cleanup
  "FavoritesManage":"Manage Favorites",
  "PicturesManage":"Manage Pictures",
  "RandomSelection":"Random Selection",
  "Randomchoice":"Random choice",
  "ofTags":"of tags",
  "YourFavorites":"Your Favorites ",
  "FavoritesAdd To":"Add To Favorites",
  "FavoritesAdd":"Add To Favorites",
  "ImageChange":"Change Image",
  "TagsEdit":"Edit Tags",
  "ThisPageTop":"Top of Page",
  "NewestImage":"Newest Images",
  "HistoryNo":"No History",
  "CommentsNo":"No Comments",
  "ImageUser":"Users Images",
  "FavoritesUser":"Users Favorites",
  "MessageSend":"Send Message",
  "RatingTimes":"Times Rated",
  "ProfileUser":"User Profile",
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

// Now, retrieve the input nodes
thenodes = document.evaluate( "//input" , document , null , XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE , null );

// Finally, perform a replacement over all input nodes
for ( var i=0 ; i<thenodes.snapshotLength ; i++ ) {
  node = thenodes.snapshotItem(i);
  // Here's the key! We must replace the "value" rather than the "data"
  s = node.value;
  for ( key in replacements ) {
    s = s.replace( regex[key] , replacements[key] );
  }
  node.value = s;
}

})();
