# 夜更かし防止アプリ「PC母ちゃん」
生活リズムを安定させて健康的な生活を送るため，PCに夢中になって夜更かししてしまうのを防止したいという動機から作成

## 機能
1. 目標就寝時刻1時間前にリマインド
1. 目標就寝時刻になったらPCを自動でスリープ＆ロック
1. どうしてもスリープさせたくない場合はお仕置きを受けることで解除できる
    - お仕置き1：指定したPCアプリを24時間実行できなくなる
    - お仕置き2：指定したWebサイトを24時間閲覧できなくなる
1. リアルお母さん的な存在の人にSlackで通知を送って監視してもらえる

## 使用言語＆ライブラリ
Python+Tkinter+ctypes+slackweb

## 使い方
通知先のSlack用WebHooks URLを設定する．
寝たい時刻をプルダウンから選択．
設定した時刻に自動的にPCがスリープ＆ロックされる．
自動スリープを解除したい場合に受けるお仕置きを設定する．
テストボタンでちゃんと動くか確認できる．
