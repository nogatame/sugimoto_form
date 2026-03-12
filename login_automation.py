import time
import pyotp
from playwright.sync_api import sync_playwright
import os

#ログイン設定。各々のデータに書き換えてね。僕はセキュリティの関係上、githubにパスワードを上げております。


MAIL_OMU = "sh23701v@st.omu.ac.jp"
OMUID = "sh23701v"
PASSWORD = os.environ.get("LONG_PASSWORD")
TOTP_SECRET = os.environ.get("TOTP_OMU")

# ------
FORM_URL = "https://forms.office.com/pages/responsepage.aspx?id=6Zxr3Lqm0069Z5-eOshc8TtS7jyhIYZOiMld3H8SQqdUMUlET1RFSTY4T0NTVlMyN0hNWjJaRVlWUSQlQCN0PWcu&origin=lprLink&route=shorturl"
#フォーム入力に使う情報
CERCLE = "落語研究会"
NAME = "栗本隆雅"
PHONE = "08038098388"
EMAIL = "kurimotoishere@gmail.com"
SUGIMOTO_ROOM = "和室、多目的室"


def login_and_open_form():
    with sync_playwright() as p:
        # ブラウザを起動 (headless=False で動作が見えるようにします)
        print("ブラウザを起動中...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(locale="ja-JP")
        page = context.new_page()

        print(f"フォームにアクセス中: {FORM_URL}")
        page.goto(FORM_URL)

        # 1. Microsoft ログイン
        print("Microsoft ログイン中...")
        page.wait_for_selector('input[type="email"]')
        page.fill('input[type="email"]', MAIL_OMU)
        page.click('input[type="submit"]')

        # 2. OMU 認証システム
        print("OMU 認証システムでログイン中...")
        # ユーザー名とパスワードの入力
        try:
            page.wait_for_selector('input[name="SM_UID"]', timeout=20000)
            page.fill('input[name="SM_UID"]', OMUID)
            page.fill('input[name="SM_PWD"]', PASSWORD)
            time.sleep(0.5)
            page.click('input[type="submit"]')
        except Exception as e:
            print(f"OMUログイン画面の待機中にエラーが発生しました: {e}")
            # 別のセレクタを試行
            page.fill('input[id*="user"]', OMUID)
            page.fill('input[id*="pass"]', PASSWORD)
            page.click('button:has-text("ログイン")')

        # 3. ワンタイムパスワード
        print("ワンタイムパスワードを生成・入力中...")
        totp = pyotp.TOTP(TOTP_SECRET)
        token = totp.now()
        print(f"現在のトークン: {token}")

        try:
            # OTP入力フィールドを待機
            page.wait_for_selector('input[name*="SM_PWD"]', timeout=15000)
            page.fill('input[name*="SM_PWD"]', token)
            page.click('input[type="submit"]')
        except Exception as e:
            print(f"OTP入力画面の待機中にエラーが発生しました: {e}")
            # 万が一フィールド名が異なる場合
            page.fill('input[type="text"]', token)
            page.keyboard.press("Enter")

        # 4. サインイン状態の維持
        try:
            page.wait_for_selector('#idSIButton9', timeout=10000)
            page.click('#idSIButton9') # 「はい」をクリック
        except:
            pass

        print("\nログインが完了しました！")
        print("このままブラウザを開いておきます。")
        # 1ページ目の操作
        for i in range(5): 
            radio = page.locator("input[type='radio']").nth(i)
            found_count = 0
            # 0.1秒だけ待って、要素が見つかるか確認（timeoutを短くするのがコツ）
            try:
                if radio.is_visible(timeout=100): 
                    radio.check()
                    print(f"{i+1}番目のラジオボタンをチェックしました")
                    found_count += 1
                    
                    # 3個見つかったらループを強制終了（ここがポイント！）
                    if found_count >= 3:
                        print("3個見つかったので検索を終了します。")
                        break
            except:
                # 見つからない場合は次のインデックスへ
                continue
        button = page.locator('button[data-automation-id="nextButton"]')
        if button.is_visible():
            print(f"「次へ」ボタンを検出。クリックします。")
            button.click()
        # 6. 2ページ目の操作
        try:
            page.wait_for_selector('input[aria-label="単一行テキスト"]', timeout=20000)
            inputs = page.locator("input").all()
            print(f"見つかった入力フィールドの数: {len(inputs)}")
            # 2. 1番目（インデックス 0）に "a" を入力
            if len(inputs) > 0:
                inputs[0].fill(CERCLE)
                inputs[1].fill(NAME)
                inputs[2].fill(PHONE)
                inputs[3].fill(EMAIL)
            button = page.locator('button[data-automation-id="nextButton"]')
            if button.is_visible():
                print(f"「次へ」ボタンを検出。クリックします。")
                button.click()
        except Exception as e:
            print(f"フォーム操作中にエラーが発生しました: {e}")
            page.screenshot(path="error_trace.png")
        # 3ページ目の操作
        try:
            page.wait_for_selector('input[type="radio"]', timeout=3000)
            radio_buttons = page.locator("input[type='radio']").all()
            print(f"見つかったラジオボタンの数: {len(radio_buttons)}")
            radio_buttons[1].check(force=True)
            # 「次へ」ボタンのクリック
            # 複数のセレクタ候補（ID, テキスト, aria-label）を試行
            
            button = page.locator('button[data-automation-id="nextButton"]')
            if button.is_visible():
                print(f"「次へ」ボタンを検出。クリックします。")
                button.click()
        except Exception as e:
            print(f"フォーム操作中にエラーが発生しました: {e}")
            page.screenshot(path="error_trace.png")
        # 4ページ目の操作
        try:
            page.wait_for_selector('input[aria-label="単一行テキスト"]', timeout=3000)
            inputs = page.locator('input').all()
            print(f"見つかった入力フィールドの数: {len(inputs)}")
            inputs[0].fill(SUGIMOTO_ROOM)
            # 「次へ」ボタンのクリック
            button = page.locator('button[data-automation-id="nextButton"]')
            if button.is_visible():
                print(f"「次へ」ボタンを検出。クリックします。")
                button.click()
        except Exception as e:
            print(f"フォーム操作中にエラーが発生しました: {e}")
            page.screenshot(path="error_trace.png")
        # 5ページ目の操作
        try:
            page.wait_for_selector('input[type="radio"]', timeout=3000)
            radio_buttons = page.locator("input[type='radio']").all()
            print(f"見つかったラジオボタンの数: {len(radio_buttons)}")
            radio_buttons[0].check(force=True)
            # 「次へ」ボタンのクリック
            button = page.locator('button[data-automation-id="nextButton"]')
            if button.is_visible():
                print(f"「次へ」ボタンを検出。クリックします。")
                button.click()
        except Exception as e:
            print(f"フォーム操作中にエラーが発生しました: {e}")
            page.screenshot(path="error_trace.png")
        # 6ページ目の操作
        try:
            page.wait_for_selector('input[type="radio"]', timeout=3000)
            radio_buttons = page.locator("input[type='radio']").all()
            print(f"見つかったラジオボタンの数: {len(radio_buttons)}")
            radio_buttons[0].check(force=True)
            # 「次へ」ボタンのクリック
            page.locator("input[type='checkbox']").check()
            button = page.locator('button[data-automation-id="submitButton"]')
            if button.is_visible():
                print(f"「次へ」ボタンを検出。クリックします。")
                button.click()
        except Exception as e:
            print(f"フォーム操作中にエラーが発生しました: {e}")
            page.screenshot(path="error_trace.png")


if __name__ == "__main__":
    login_and_open_form()
    
