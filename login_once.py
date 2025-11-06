from playwright.sync_api import sync_playwright

LOGIN_URL = "https://sso.ice.com/appUserLogin?redirectUrl=https://ica.ice.com/ICA/Login/SsoUlp&loginApp=ICA#/pageLogin"
SESSION_FILE = "ice_session.json"


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(LOGIN_URL)
        print(
            """
1. Enter your ICE email and password.
2. Complete 2FA (authenticator / SMS).
3. Wait until you're redirected to https://ica.ice.com/ICA/Main (the calculator page).
Then return to this terminal.
        """
        )
        input("Press ENTER here once you see the main ICA dashboard in the browser... ")

        # Save cookies and localStorage
        context.storage_state(path=SESSION_FILE)
        print(
            f"\n✅ Session saved to {SESSION_FILE}. Future runs won't need login or 2FA."
        )
        browser.close()


if __name__ == "__main__":
    main()
