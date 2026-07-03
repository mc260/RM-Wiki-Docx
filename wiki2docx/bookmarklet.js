"""
RoboMaster Wiki Scraper — pure browser Console script (bookmarklet).

Use this when you can't or don't want to install Python + Playwright.
Open the wiki page in Chrome (logged in), paste this into F12 Console,
wait for completion, then copy(allData) to get the JSON.

HOW TO USE:
1. Open https://bbs.robomaster.com/wiki/YOUR_WIKI_ID in Chrome (logged in)
2. F12 → Console tab
3. Paste this ENTIRE script, press Enter
4. Watch the progress — tree expands, then sections are clicked one by one
5. When you see === DONE ===, run: copy(JSON.stringify(allData))
6. Open Notepad, Ctrl+V, save as data.json
7. Run: python -m wiki2docx --json-file data.json --output output.docx
   OR ask someone who has the CLI installed to generate the Word file for you
"""

(async function run() {
  console.clear();
  console.log("=== RoboMaster Wiki Scraper ===");

  // ---- Step 1: Expand all tree nodes ----
  console.log("Phase 1: Expanding sidebar tree...");
  for (let round = 0; round < 12; round++) {
    const closed = document.querySelectorAll(".ant-tree-switcher_close");
    if (closed.length === 0) {
      console.log(`  Round ${round + 1}: all expanded!`);
      break;
    }
    console.log(`  Round ${round + 1}: expanding ${closed.length} nodes…`);
    closed.forEach((el) => el.click());
    await new Promise((r) => setTimeout(r, 1200));
  }
  // Final sweep
  document.querySelectorAll(".ant-tree-switcher_close").forEach((el) => el.click());
  await new Promise((r) => setTimeout(r, 1500));

  // ---- Step 2: Extract tree ----
  console.log("Phase 2: Extracting menu tree…");
  function extractTree(ul) {
    const items = [];
    if (!ul) return items;
    ul.querySelectorAll(":scope > li").forEach((li) => {
      const box = li.querySelector(".menuAction__box");
      const title = box ? box.getAttribute("title") || box.innerText.trim() : "";
      const child = li.querySelector(":scope > ul");
      items.push({ title, children: child ? extractTree(child) : [] });
    });
    return items;
  }
  const root = document.querySelector(".wikiMenu__tree .ant-tree");
  const menu = root ? extractTree(root) : [];

  const sections = [];
  function flatten(items, level) {
    items.forEach((item) => {
      const isLeaf = item.children.length === 0;
      sections.push({ title: item.title, level, isLeaf, content_html: "", content: "" });
      if (item.children.length) flatten(item.children, level + 1);
    });
  }
  flatten(menu, 0);
  console.log(`  ${sections.length} nodes (${sections.filter((s) => s.isLeaf).length} leaves)`);

  // ---- Step 3: Click-scrape ----
  console.log(`Phase 3: Scraping ${sections.length} sections…`);

  let success = 0;
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i];
    const label = `[${i + 1}/${sections.length}]`;

    let clicked = false;
    for (const box of document.querySelectorAll(".menuAction__box")) {
      if ((box.getAttribute("title") || box.innerText.trim()) === section.title) {
        box.click();
        clicked = true;
        break;
      }
    }
    if (!clicked) {
      console.log(`${label} ${section.title} — NOT FOUND`);
      continue;
    }

    await new Promise((r) => setTimeout(r, 600));

    const container = document.getElementById("rick-viewer-container");
    if (container && !container.querySelector(".client-only-placeholder")) {
      const editor = container.querySelector("[data-slate-editor], .editor-content-container");
      if (editor) {
        section.content_html = editor.innerHTML;
        section.content = editor.innerText || "";
      } else {
        section.content_html = container.innerHTML;
        section.content = container.innerText || "";
      }
    }

    if (section.content) {
      success++;
      const preview = section.content.substring(0, 60).replace(/\n/g, " ");
      console.log(`${label} ${"  ".repeat(section.level)}${section.title} [${section.content.length} chars] "${preview}…"`);
    } else {
      console.log(`${label} ${"  ".repeat(section.level)}${section.title} — (empty)`);
    }
  }

  // ---- Done ----
  window.allData = {
    wiki_url: location.href,
    title: document.title,
    scraped_at: new Date().toISOString(),
    total_nodes: sections.length,
    leaf_nodes: sections.filter((s) => s.isLeaf).length,
    with_content: success,
    empty: sections.length - success,
    sections,
  };

  console.log("\n=== DONE ===");
  console.log(`Nodes: ${sections.length} | Content: ${success} | Empty: ${sections.length - success}`);
  console.log("");
  console.log("▶ Run this to copy the data:");
  console.log("  copy(JSON.stringify(allData))");
  console.log("▶ Then paste into a file called data.json");
  console.log("▶ Then run: python -m wiki2docx --json-file data.json --output output.docx");
})();
