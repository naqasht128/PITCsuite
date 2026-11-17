• **Ultra-Fast Launch Performance:** Switched to native onedir architecture and asynchronous licensing checks, reducing startup time from 20+ seconds down to 1–2 seconds.
• **Massive Size Reduction:** Excluded unneeded OpenCV binary dependencies and trimmed unnecessary models, cutting installer and executable size by over 130 MB.
• **Direct Update Flow:** Clicking 'Update Available' now opens the update prompt and release notes directly, skipping intermediate dialogs.
• **Batch Progress Logging:** Real-time progress sequence (e.g. 1/3: Processing <AC No>) now displays cleanly in Bill Scrapper and E-Bills Downloader logs.
• **Automatic CAPTCHA Input:** Integrated automatic CAPTCHA recognition and input across portal bill scraping and downloader workflows.
• **Chrome Session Stability:** Suppressed crash recovery and 'Restore pages' prompts when Chrome closes and restarts during batch runs.
• **Modern PDF Engine:** Upgraded to modern PyMuPDF API across tools for clean, deprecation-free rendering and stamping.
• **Thread Safety & Shutdown:** Resolved QThread background runner destruction on application exit for clean shutdowns.