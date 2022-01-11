document.getElementById('getWordForm').addEventListener('submit', resolve);

function resolve(evt) {
	evt.preventDefault();
	const username = this.elements[0].value;
	const article = this.elements[1].value;
	if (username.length === 0) {
		return;
	}
	this.reset();
	try {
		fetch("https://api.rss2json.com/v1/api.json?rss_url=https://medium.com/feed/@" + username)
			.then(result => {
				result.json().then(async data => {
					let allPublishes = data.items;
					let posts = allPublishes.filter(item => item.categories.length > 0);
					if (article.length > 0) {
						posts = posts.filter(item => item.title.toLowerCase() === article.toLowerCase());
					}
					
					for (let i = 0; i < 1; i++) { // posts.length
						const title = posts[i].title;
						
						// set up document
						let doc = new docx.Document({
							'title': title,
							'styles': myStyles(),
							'numbering': {
								'config': [
									{
										reference: 'medium-numbering',
										levels: [
											{
												level: 0,
												format: "decimal",
												text: "%1.",
												alignment: docx.AlignmentType.START,
												style: {
													run: {
														font: {
															name: "Calibri",
														},
														size: 24,
													},
													paragraph: {
														indent: {left: 720, hanging: 260},
													},
												},
											},
										],
									},
								],
							},
						});
						// retrieve and set up HTML
						let textContent = posts[i].description;
						let body = document.createElement('body');
						let header = document.createElement('h1');
						header.innerText = title;
						body.innerHTML = header.outerHTML + textContent;
						
						let paras = [];
						
						// parse into document pieces
						let nodes = body.children; // elements only!
						for (let n = 0; n < nodes.length; n++) {
							if (nodes[n].nodeName === "OL" || nodes[n].nodeName === "UL") {
								let listItems = nodes[n].children; // not text nodes like line breaks
								for (let subN = 0; subN < listItems.length; subN++) {
									let liText = await getTextRuns(listItems[subN], ['_li'], doc);
									let liPara;
									if (nodes[n].nodeName === "OL") {
										liPara = new docx.Paragraph({
											'children': liText,
											'style': 'myNormal',
											numbering: {
												reference: 'medium-numbering',
												level: 0,
											},
										});
									} else {
										liPara = new docx.Paragraph({
											'children': liText,
											'style': 'myNormal',
											bullet: {
												level: 0,
											},
										});
									}
									paras.push(liPara);
								}
								continue;
							}
							let style = getWordStyle(nodes[n].nodeName);
							let textRuns = await getTextRuns(nodes[n], [style], doc);
							let para;
							if (textRuns.length === 0 && style === 'myHr') {
								para = new docx.Paragraph({
									'style': style,
									'children': [new docx.TextRun('* * *')],
								});
							} else {
								para = new docx.Paragraph({
									'style': style,
									'children': textRuns,
								});
							}
							paras.push(para);
						}
						
						doc.addSection({
							'children': paras,
						});
						
						// save
						docx.Packer.toBlob(doc).then(blob => {
							let filename = title + ".docx";
							saveAs(blob, filename);
						});
						
					}
				})
			});
	} catch(err) {
		console.log(err);
	}
}
	
function saveDocumentToFile(doc, filename) {
	const pckr = new docx.Packer();
	pckr.toBlob(doc).then(blob => {
		let docBlob = blob.slice(0, blob.size, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
		
		saveAs(docBlob, filename);
	});
}

function getWordStyle(nodeName) {
	let style;
	switch(nodeName.toUpperCase()) {
		case "H1":
			style = "myTitle";
			break;
		case "H3":
			style = "myHeading1";
			break;
		case "H4":
			style = "myHeading2";
			break;
		case "BLOCKQUOTE":
			style = "myQuote";
			break;
		case "STRONG":
			style = "myStrong";
			break;
		case "EM":
			style = "myEmphasis";
			break;
		case "FIGCAPTION":
		case "FIGURE":
			style = "myCaption";
			break;
		case "HR":
			style = "myHr";
			break;
		default: // includes "P", "A", "BR", "IMG", "LI", "OL", "UL"
			style = "myNormal";
			break;
	}
	return style;
}

async function getImage(dimensions, href, imageSizer) {
	return new Promise(async (resolve) => {
		if (dimensions.error) {
			console.log("Images are probably rate-limited. Wait ~1 hour then try again.");
			let img = document.createElement('img');
			img.onload(async () => {
				dimensions = { height: img.clientHeight, width: img.clientWidth };
				let aspectRatio = dimensions.height / dimensions.width;
				let width = inchesToPixels(6.2);
				let height = width * aspectRatio;
				let blob = await fetch(href).then(result => result.blob());
				resolve([blob, width, height]);
			});
			img.src = href;
			while (imageSizer.childElementCount) imageSizer.removeChild(imageSizer.firstChild);
			imageSizer.appendChild(img);
		} else {
			let aspectRatio = dimensions.height / dimensions.width;
			let width = inchesToPixels(6.2);
			let height = width * aspectRatio;
			let blob = await fetch(href).then(result => result.blob());
			resolve([blob, width, height]);
		}
	})
}

async function getTextRuns(node, parentStyleList, doc) {
	let output = [];
	let kids = node.childNodes;
	const imageSizer = document.querySelector('.imageSizer');
	for (let i = 0; i < kids.length; i++) { // now we want all children
		let run = null;
		if (kids[i].nodeType === Node.ELEMENT_NODE) {
			if (kids[i].nodeName.toUpperCase() === "IMG") {
				let href = kids[i].getAttribute('src');
				let dimensions = await fetch('https://api.saasify.sh/1/call/dev/image-size/sizebyurl?url=' + href, {
					headers: {
						'Content-Type': 'application/json',
						'Authorization': myToken,
					},
				}).then(result => result.json());
				let [blob, width, height] = await getImage(dimensions, href, imageSizer);
				run = docx.Media.addImage(doc, blob, width, height);
			} else {
				if (kids[i].nodeName.toUpperCase() === "A") {
					let href = kids[i].getAttribute('href');
					if (href.length > 0) {
						run = doc.createHyperlink(href, kids[i].textContent);
					}
				}
				if (parentStyleList.includes('_li') && !parentStyleList.includes('myNormal')) {
					style = 'myNormal';
				}
				if (run === null) {
					let style = getWordStyle(kids[i].nodeName);
					let myOwnKids = await getTextRuns(kids[i], Array.prototype.concat(parentStyleList, [style]), doc);
					let properties = { 'children': myOwnKids, 'style': style };
					if (parentStyleList.includes('myStrong')) {
						properties.bold = true;
					}
					if (style === 'myCaption') {
						properties.break = 1;
					}
					run = new docx.TextRun(properties);
				}
			}
		} else {
			let properties = { 'text': kids[i].textContent };
			if (parentStyleList.includes('myStrong')) {
				properties.bold = true;
			}
			if (parentStyleList.includes('_li') && !parentStyleList.includes('myNormal')) {
				properties.font = 'Calibri';
				properties.size = 24;
			}
			run = new docx.TextRun(properties);
		}
		output.push(run);
	}
	return output;
}

function myStyles() {
	let halfInch = docx.convertInchesToTwip(0.5);
	return {
		default: {
			listNumbering: {
				run: {
					font: {
						name: "Calibri",
					},
					size: 24,
				},
			},
		},
		paragraphStyles: [
			{
				id: 'myNormal',
				name: 'MediumNormal',
				basedOn: 'Normal',
				run: {
					font: {
						name: "Calibri",
					},
					size: 24,
				},
				paragraph: {
					spacing: {
						before: 120,
					},
				},
			},
			{
				id: 'myTitle',
				name: 'MediumTitle',
				basedOn: 'Title',
				run: {
					size: 42,
					bold: true,
					font: {
						name: "Cambria",
					},
				},
				paragraph: {
					spacing: {
						after: 240,
					},
				},
			},
			{
				id: 'myHeading1',
				name: 'MediumHeading1',
				basedOn: 'Heading1',
				run: {
					size: 36,
					bold: true,
				},
				paragraph: {
					spacing: {
						before: 240,
					},
				},
			},
			{
				id: 'myHeading2',
				name: 'MediumHeading2',
				basedOn: 'Heading2',
				run: {
					size: 28,
					italics: true,
				},
				paragraph: {
					spacing: {
						before: 120,
					},
				},
			},
			{
				id: 'myQuote',
				name: 'MediumQuote',
				basedOn: 'myNormal',
				run: {
					italics: true,
				},
				paragraph: {
					indent: {
						right: halfInch,
						left: halfInch,
					},
					spacing: {
						before: 120,
						after: 120,
					},
				},
			},
			{
				id: 'myCaption',
				name: 'MediumCaption',
				basedOn: 'myNormal',
				paragraph: {
					spacing: {
						after: 0,
					},
					alignment: docx.AlignmentType.CENTER,
				}
			},
			{
				id: 'myHr',
				name: 'MediumHorizRule',
				basedOn: 'myNormal',
				paragraph: {
					alignment: docx.AlignmentType.CENTER,
				}
			},
		],
		characterStyles: [
			{
				id: 'myEmphasis',
				name: 'MediumEmphasis',
				basedOn: 'myNormal',
				run: {
					italics: true,
				},
			},
			{
				id: 'myStrong',
				name: 'MediumStrong',
				basedOn: 'myNormal',
				run: {
					bold: true,
				},
			},
		],
	}
}

function inchesToPixels(inches) {
	return twipsToPixels(docx.convertInchesToTwip(inches));
}

function twipsToPixels(twips) {
	return twips / 15;
}