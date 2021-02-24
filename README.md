# medium-to-word
Takes a Medium article and makes a Word document out of it.

## Vanilla vs Node.js
This started as a quick project to handle some article backups for a Medium publication I work with. I used vanilla JS, which (most significantly) means using the free public API for [image-size](https://image-size.saasify.sh/docs#section/Quick-Start), which is rate-limited to 15 queries/hour. If you want to do the minor changes needed to revamp this for Node.js, you can just install and import [image-size](https://github.com/image-size/image-size) yourself. If you do that, you can also speed up the Word document generation by installing and importing [docx](https://github.com/dolanmiu/docx) instead of using the public API.

## Usage
Otherwise, you just need a place to put these files (http://localhost/ will even do it) and a [token from Saasify](https://image-size.saasify.sh/pricing). Put the Saasify token in a tokens.js file (as seen in the tokens.js.example file in the repo). Then open the webpage and download away.

## FAQ
### Why not use Medium's API?
[Medium's API](https://github.com/Medium/medium-api-docs) covers writing articles and uploading images. The only things you can GET are a user's publications and a publication's users. That doesn't help us, so we use Medium's RSS feeds and convert them from XML to JSON for consumption.

### Why not just consume the XML?
I was in a hurry and consuming JSON sounded easier.

### Where are my article's horizontal rules / asterisk section breaks?
Medium doesn't include those in the RSS feed. We would handle them if they did.

### I don't like the formatting in my generated Word document.
Fork or change it after downloading.

### No, it's literally broken.
[Make an issue](https://github.com/VersorVerbi/medium-to-word/issues/new/choose), I guess?
