(() => {
  const data = window.__CAROUSEL_DATA__;
  document.documentElement.dataset.theme = data.meta.theme || "classic";
  const root = document.getElementById("carousel");
  const pageRecords = [];
  const BRAND_NAME = "智富界";
  const BRAND_TAGLINE = "看懂AI，用好AI，投资AI";

  function element(tagName, className, html) {
    const node = document.createElement(tagName);
    if (className) node.className = className;
    if (html !== undefined) node.innerHTML = html;
    return node;
  }

  function textElement(tagName, className, text) {
    const node = element(tagName, className);
    node.textContent = text || "";
    return node;
  }

  function createBrand() {
    const brand = element("div", "brand-lockup");
    brand.appendChild(textElement("strong", "brand-name", BRAND_NAME));
    return brand;
  }

  function createFooter() {
    const footer = element("footer", "page-footer");
    footer.appendChild(createBrand());
    footer.appendChild(textElement("span", "tagline", BRAND_TAGLINE));
    footer.appendChild(textElement("span", "page-number", ""));
    return footer;
  }

  function createCover() {
    const page = element("section", "page-card cover-page");
    page.dataset.kind = "cover";
    page.appendChild(textElement("div", "cover-kicker", data.meta.kicker));
    const main = element("div", "cover-main");
    main.appendChild(element("h1", "cover-title", data.meta.titleHtml));
    if (data.meta.subtitleHtml) main.appendChild(element("p", "cover-subtitle", data.meta.subtitleHtml));
    page.appendChild(main);
    const bottom = element("div", "cover-bottom");
    bottom.appendChild(createBrand());
    bottom.appendChild(textElement("span", "cover-tagline", BRAND_TAGLINE));
    bottom.appendChild(textElement("span", "page-number", ""));
    page.appendChild(bottom);
    root.appendChild(page);
    pageRecords.push({ page, flow: null });
  }

  function createBodyPage() {
    const page = element("section", "page-card body-page");
    page.dataset.kind = "body";
    const flow = element("main", "page-flow");
    page.appendChild(flow);
    page.appendChild(createFooter());
    root.appendChild(page);
    const record = { page, flow };
    pageRecords.push(record);
    return record;
  }

  function renderList(block, nested = false) {
    const list = element(block.ordered ? "ol" : "ul", nested ? "nested-list" : "content-block list-block");
    block.items.forEach((item) => {
      const listItem = element("li", item.task ? "task-list-item" : "");
      if (item.task) {
        const checkbox = textElement("span", `task-checkbox${item.checked ? " is-checked" : ""}`, item.checked ? "✓" : "");
        checkbox.setAttribute("aria-hidden", "true");
        listItem.appendChild(checkbox);
      }
      listItem.appendChild(element("span", "list-item-text", item.html));
      item.children?.forEach((child) => listItem.appendChild(renderList(child, true)));
      list.appendChild(listItem);
    });
    return list;
  }

  function renderBlock(block) {
    if (block.type === "paragraph") return element("p", "content-block body-paragraph", block.html);
    if (block.type === "heading") return element(`h${Math.min(6, block.level + 1)}`, `content-block subheading heading-level-${block.level}`, block.html);
    if (block.type === "section") {
      const section = element("div", "content-block block-section");
      section.appendChild(element("span", "section-diamond"));
      section.appendChild(element("strong", "section-text", block.html));
      return section;
    }
    if (block.type === "metrics") {
      const metrics = element("dl", "content-block metrics");
      block.items.forEach((item) => {
        const row = element("div", "metric-row");
        row.appendChild(element("dt", "metric-label", item.labelHtml));
        row.appendChild(element("dd", "metric-value", item.valueHtml));
        metrics.appendChild(row);
      });
      return metrics;
    }
    if (block.type === "marker") {
      const marker = element("p", "content-block marker-block");
      marker.appendChild(element("span", "marker-ink", block.html));
      return marker;
    }
    if (block.type === "callout") return element("div", "content-block callout-block", block.html);
    if (block.type === "lead") return element("p", "content-block lead-block", block.html);
    if (block.type === "source") return element("p", "content-block source-block", block.html);
    if (block.type === "quote") return element("blockquote", "content-block quote-block", block.html);
    if (block.type === "list") return renderList(block);
    if (block.type === "hr") return element("hr", "content-block markdown-rule");
    if (block.type === "code") {
      const wrapper = element("div", "content-block code-block");
      if (block.language) wrapper.appendChild(textElement("div", "code-language", block.language));
      const pre = element("pre", "");
      const code = textElement("code", "", block.raw);
      pre.appendChild(code);
      wrapper.appendChild(pre);
      return wrapper;
    }
    if (block.type === "image") {
      const figure = element("figure", "content-block markdown-image");
      const image = element("img", "");
      image.src = block.src;
      image.alt = block.alt || "";
      if (block.title) image.title = block.title;
      figure.appendChild(image);
      if (block.captionHtml) figure.appendChild(element("figcaption", "", block.captionHtml));
      return figure;
    }
    if (block.type === "table") {
      const wrapper = element("div", "content-block markdown-table-wrap");
      const table = element("table", "markdown-table");
      const head = element("thead");
      const headRow = element("tr");
      block.headers.forEach((cell, index) => {
        const header = element("th", "", cell.html);
        header.style.textAlign = block.alignments[index] || "left";
        headRow.appendChild(header);
      });
      head.appendChild(headRow);
      table.appendChild(head);
      const body = element("tbody");
      block.rows.forEach((row) => {
        const tableRow = element("tr");
        row.forEach((cell, index) => {
          const dataCell = element("td", "", cell.html);
          dataCell.style.textAlign = block.alignments[index] || "left";
          tableRow.appendChild(dataCell);
        });
        body.appendChild(tableRow);
      });
      table.appendChild(body);
      wrapper.appendChild(table);
      return wrapper;
    }
    if (block.type === "footnotes") {
      const footnotes = element("section", "content-block footnotes-block");
      footnotes.appendChild(textElement("div", "footnotes-title", "注释"));
      const list = element("ol");
      block.items.forEach((item) => {
        const listItem = element("li", "", item.html);
        listItem.value = item.number;
        list.appendChild(listItem);
      });
      footnotes.appendChild(list);
      return footnotes;
    }
    if (block.type === "html") return element("div", "content-block html-block", block.html);
    throw new Error(`Unsupported rendered block: ${block.type}`);
  }

  function overflows(record) {
    return record.flow.scrollHeight > record.flow.clientHeight + 1 || record.flow.scrollWidth > record.flow.clientWidth + 1;
  }

  function appendWithPagination(block, current) {
    const node = renderBlock(block);
    current.flow.appendChild(node);
    if (!overflows(current)) return current;

    current.flow.removeChild(node);
    const previous = current.flow.lastElementChild;
    const carry = [];
    if (previous?.classList.contains("block-section")) {
      carry.push(previous);
    } else if (node.classList.contains("source-block") && previous) {
      carry.unshift(previous);
      const beforePrevious = previous.previousElementSibling;
      if (beforePrevious?.classList.contains("marker-block")) carry.unshift(beforePrevious);
    }
    carry.forEach((carried) => current.flow.removeChild(carried));

    const next = createBodyPage();
    carry.forEach((carried) => next.flow.appendChild(carried));
    next.flow.appendChild(node);
    return next;
  }

  if (data.meta.cover) createCover();
  let current = createBodyPage();
  data.blocks.forEach((block) => {
    if (block.type === "pagebreak") {
      if (current.flow.children.length) current = createBodyPage();
      return;
    }
    current = appendWithPagination(block, current);
  });

  if (!current.flow.children.length && pageRecords.filter((record) => record.flow).length > 1) {
    current.page.remove();
    pageRecords.splice(pageRecords.indexOf(current), 1);
  }

  pageRecords.forEach((record, index) => {
    const pageNumber = record.page.querySelector(".page-number");
    if (pageNumber) pageNumber.textContent = String(index + 1).padStart(2, "0");
  });

  window.__renderReport = {
    pageCount: pageRecords.length,
    overflowPages: pageRecords.map((record, index) => record.flow && overflows(record) ? index + 1 : null).filter(Boolean)
  };
  document.body.dataset.renderReady = "true";
})();
