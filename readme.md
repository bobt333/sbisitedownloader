# SBI Site Downloader

SBI/SiteSell is awesome. I highly recommend them. But it is a proprietary system and migrating files out to another platform, like WordPress, is NOT EASY.

If you use "UYOH" (Upload Your Own HTML) you likely have all the primary files for your site already stored locally, but if you have many C2 posts, you have no backup of those to migrate to anohter system. And if you use BB (Block Builder) instead of UYOH, you may not even have your core content saved on your local drive.

My "SBI Site Downloader" (sbisitedownloader.vbs) is a relatively simple VB Script app that downloads all files listed in your site's sitemap file and saves them locally.

NOTE: It ONLY downloads the web files (HTML), not special files, support files, images, etc. It is really intended to salvage your core textual content, not all layout/etc features. That's something you'll need to work out on the WordPress side.

**My strongest advice: Read this whole document slowly and carefully all the way through at least once before you do anything. Just read without actually doing anything else. Then, when you start this process, do it slowly, one step at a time, reading carefully... If you have built an SBI/SiteSell site following the Action Guide, I know you're up to it.**

This is open source. Download it. Change it. Whatever you want. It's only intended for Windows (sorry Mac/Linux). NO GAURANTEES. USE AT YOUR OWN RISK. But it's very simple code written in VB Script that nearly any programmer should be able to validate.


# How to use it:

1) Create a folder. Name it anything, but for example, "C:\sitedownloader"

2) From github download sbisitedownloader.vbs to that folder. You can save this readme.md file there too if you like, but if you're reading this you don't need to.

3) Find your site's sitemap address and paste it into sbisitedownloader.vbs...

    a) browse to mysite.com/robots.txt (of course replacing "mysite.com" with your domain name). One of the first lines of that file should be something like "Sitemap: https://www.mysite.com/OmrvZpiX.xml".

    b) Copy that address ("https://www.mysite.com/OmrvZpiX.xml").

    c) Right-click sbisitedownloader.vbs and select "Open with Notepad" (OR, open Notepad, then from the menu, select "File", then "Open", then use the file dialog to locate sbisitedownloader.vbs).

    d) Edit the FIRST LINE ONLY. Paste your sitemap address copied in step "b" above between the double quotes.

        strXMLURL = "https://www.mysite.com/OmrvZpiX.xml"

4) Double-click sbisitedownloader.vbs.

5) NOW JUST WAIT. I'm sorry that it doesn't display an hourglass or anything like that. On my computer this takes 2-3 minutes to download 400 files. My computer is fairly up-to-date (2022, AMD Ryzen 9, 16 GB RAM, 1 TB SSD). If your computer is older, weaker, slower, be patient. At the end of the process you should see a pop-up message "Finished. Downloaded 409 files to..." This process will create a folder named "files" inside that folder you created in step 1. All the files will be there.
