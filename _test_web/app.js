(function () {
  "use strict";

  // ---------- 工具 ----------
  function $(id) { return document.getElementById(id); }
  function norm(s) { return (s || "").replace(/\u00a0/g, " ").replace(/\u3000/g, " ").replace(/\s+/g, " ").trim(); }
  function padTime(t) {
    if (!t) return "";
    var m = t.match(/^(\d{1,2}):(\d{2})$/);
    return m ? (m[1].padStart(2, "0") + ":" + m[2]) : t;
  }
  function esc(s) {
    if (s == null) return "";
    return String(s).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\r/g, "").replace(/\n/g, "\\n");
  }

  // ---------- zip 解压（仅取 word/document.xml） ----------
  async function extractDocumentXml(buffer) {
    var dv = new DataView(buffer);
    var len = buffer.byteLength;
    var eocd = -1;
    for (var i = len - 22; i >= len - 22 - 65536 && i >= 0; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("不是有效的 .docx（zip 结构不正确）。");
    var count = dv.getUint16(eocd + 10, true);
    var p = dv.getUint32(eocd + 16, true);
    for (var j = 0; j < count; j++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var method = dv.getUint16(p + 10, true);
      var compSize = dv.getUint32(p + 20, true);
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var localOff = dv.getUint32(p + 42, true);
      var nameBytes = new Uint8Array(buffer, p + 46, nameLen);
      var name = new TextDecoder("utf-8").decode(nameBytes);
      if (name === "word/document.xml") {
        var lNameLen = dv.getUint16(localOff + 26, true);
        var lExtraLen = dv.getUint16(localOff + 28, true);
        var dataStart = localOff + 30 + lNameLen + lExtraLen;
        var comp = new Uint8Array(buffer, dataStart, compSize);
        if (method === 0) return new TextDecoder("utf-8").decode(comp);
        if (method === 8) {
          if (typeof DecompressionStream === "undefined")
            throw new Error("当前浏览器不支持解压，请使用较新的 Chrome / Edge / Firefox / Safari。");
          var ds = new DecompressionStream("deflate-raw");
          var stream = new Blob([comp]).stream().pipeThrough(ds);
          var out = await new Response(stream).arrayBuffer();
          return new TextDecoder("utf-8").decode(out);
        }
        throw new Error("不支持的压缩方式：" + method);
      }
      p += 46 + nameLen + extraLen + commentLen;
    }
    throw new Error("文档里没有 word/document.xml。");
  }

  // ---------- XML → 表格矩阵 ----------
  function decodeEntities(s) {
    return s
      .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 10)); })
      .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  }
  function parseMatrix(xmlText) {
    // 用正则提取 w:tr / w:tc，并展开 gridSpan（不依赖 DOMParser 的命名空间匹配）
    var trRe = /<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g;
    var matrix = [];
    var tr;
    while ((tr = trRe.exec(xmlText))) {
      var row = [];
      var tcRe = /<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g;
      var tc;
      while ((tc = tcRe.exec(tr[0]))) {
        var span = 1;
        var gs = tc[0].match(/<w:gridSpan\b[^>]*>/);
        if (gs) {
          var gv = gs[0].match(/w:val="(\d+)"/);
          if (gv) span = parseInt(gv[1], 10) || 1;
        }
        var tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
        var txt = "";
        var tm;
        while ((tm = tRe.exec(tc[0]))) txt += decodeEntities(tm[1]);
        var value = norm(txt);
        for (var k = 0; k < span; k++) row.push(value);
      }
      matrix.push(row);
    }
    for (var mr = 0; mr < matrix.length; mr++) {
      for (var mc = 0; mc < matrix[mr].length; mc++) {
        if (matrix[mr][mc].indexOf("星期一") >= 0) return matrix;
      }
    }
    throw new Error("没有找到含“星期一…星期日”表头的课表。");
  }

  // ---------- 单元格 → 课程 ----------
  function extractLocation(after) {
    if (!after) return "";
    var found = [];
    var re = /([A-Z]\d+-\d+(?:[^\s]*)?)|(体育场[^\s]*)|([^\s]*教室)/g;
    var m;
    while ((m = re.exec(after))) {
      var v = m[0].trim();
      if (v && found.indexOf(v) < 0) found.push(v);
    }
    if (found.length) return found.join(" / ");
    var m2 = after.match(/校区\s*([^\s]+)/);
    return m2 ? m2[1].trim() : "";
  }
  function extractTeacher(after, loc) {
    if (!after) return "（未注明）";
    var clean = after;
    if (loc) loc.split("/").forEach(function (p) { if (p.trim()) clean = clean.replace(p.trim(), " "); });
    var names = [];
    var re = /([A-Za-z][A-Za-z.,]*(\s[A-Za-z.,]+)*|[\u4e00-\u9fa5]+)\s*\([A-Za-z0-9]+\)/g;
    var m;
    while ((m = re.exec(clean))) {
      var n = m[1].trim();
      if (n && names.indexOf(n) < 0) names.push(n);
    }
    return names.length ? names.join("/") : "（未注明）";
  }
  function parseCell(cell, weekday) {
    var t = norm(cell);
    if (!t) return [];
    if (t.indexOf("教学班代码") < 0) {
      return [{ name: t, weekday: weekday, ws: 1, we: 17, st: "", et: "", loc: "", teacher: "（未注明）", code: "" }];
    }
    var parts = t.split("教学班代码");
    var nextName = parts[0].trim();
    var list = [];
    for (var i = 1; i < parts.length; i++) {
      var seg = parts[i];
      var code = "", tailName = "", ws = 0, we = 0, t1 = "", t2 = "", after = "";
      var m = seg.match(/^[：:]?\s*([A-Za-z0-9_#]+)\s*\((\d{1,2})\s*[~\-]\s*(\d{1,2})周\)\s*\((\d+)-(\d+)节\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\)\s*(.*?)(?=人数|$)/);
      if (m) {
        code = m[1]; ws = parseInt(m[2], 10) || 0; we = parseInt(m[3], 10) || 0;
        t1 = m[6]; t2 = m[7]; after = m[8];
        var qi = seg.indexOf("人数");
        if (qi >= 0) {
          var nm = seg.slice(qi).match(/人数\s*[：:]\s*\d+\/\d+(.*)$/);
          if (nm) tailName = nm[1].trim();
        }
      }
      var loc = extractLocation(after);
      list.push({
        name: nextName, code: code, weekday: weekday,
        ws: ws > 0 ? ws : 1, we: we > 0 ? we : 17,
        st: padTime(t1), et: padTime(t2),
        loc: loc, teacher: extractTeacher(after, loc)
      });
      nextName = tailName;
    }
    return list;
  }
  function parseSchedule(matrix) {
    var names = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
    var header = -1;
    for (var i = 0; i < matrix.length; i++) {
      for (var c = 0; c < matrix[i].length; c++) if (matrix[i][c].indexOf("星期一") >= 0) { header = i; break; }
      if (header >= 0) break;
    }
    var cols = [];
    for (var col = 0; col < matrix[header].length; col++) {
      for (var w = 0; w < names.length; w++) {
        if (matrix[header][col].indexOf(names[w]) >= 0) { cols.push({ col: col, weekday: w + 1 }); break; }
      }
    }
    var result = [];
    for (var r = 0; r < matrix.length; r++) {
      for (var q = 0; q < cols.length; q++) {
        var cell = (r < matrix.length && cols[q].col < matrix[r].length) ? matrix[r][cols[q].col] : "";
        if (!cell || cell.indexOf("星期") >= 0) continue;
        if (cell.indexOf("教学班代码") < 0 && /^第?\d+节/.test(cell)) continue;
        result = result.concat(parseCell(cell, cols[q].weekday));
      }
    }
    return result;
  }

  // ---------- ICS 生成 ----------
  function fmtDate(d) {
    function p2(n) { return (n < 10 ? "0" : "") + n; }
    return d.getFullYear() + p2(d.getMonth() + 1) + p2(d.getDate());
  }
  function buildIcs(entries, firstMon, remind, calName, withTz) {
    var L = [];
    L.push("BEGIN:VCALENDAR");
    L.push("VERSION:2.0");
    L.push("PRODID:-//KejianToIcs//Course Calendar//CN");
    L.push("CALSCALE:GREGORIAN");
    L.push("METHOD:PUBLISH");
    L.push("X-WR-CALNAME:" + esc(calName || "课表"));
    if (withTz) {
      L.push("X-WR-TIMEZONE:Asia/Shanghai");
      L.push("BEGIN:VTIMEZONE");
      L.push("TZID:Asia/Shanghai");
      L.push("BEGIN:STANDARD");
      L.push("DTSTART:19700101T000000");
      L.push("TZOFFSETFROM:+0800");
      L.push("TZOFFSETTO:+0800");
      L.push("TZNAME:CST");
      L.push("END:STANDARD");
      L.push("END:VTIMEZONE");
    }
    var dtstamp = "";
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (!e.name) continue;
      var day = new Date(firstMon.getFullYear(), firstMon.getMonth(), firstMon.getDate());
      day.setDate(day.getDate() + (e.ws - 1) * 7 + (e.weekday - 1));
      var ymd = fmtDate(day);
      var t1 = e.st.replace(":", "") + "00";
      var t2 = e.et.replace(":", "") + "00";
      var dstart = withTz ? "DTSTART;TZID=Asia/Shanghai:" + ymd + "T" + t1 : "DTSTART:" + ymd + "T" + t1;
      var dend = withTz ? "DTEND;TZID=Asia/Shanghai:" + ymd + "T" + t2 : "DTEND:" + ymd + "T" + t2;
      var desc = "教学班代码：" + e.code + "\n教师：" + e.teacher + "\n周次：" + e.ws + "~" + e.we + "周";
      L.push("BEGIN:VEVENT");
      L.push("UID:" + e.code + "-" + (i + 1) + "@kejiantoics.local");
      L.push("DTSTAMP:" + dtstamp);
      L.push(dstart);
      L.push(dend);
      L.push("RRULE:FREQ=WEEKLY;COUNT=" + (e.we - e.ws + 1));
      L.push("SUMMARY:" + esc(e.name));
      L.push("LOCATION:" + esc(e.loc));
      L.push("DESCRIPTION:" + esc(desc));
      if (remind > 0) {
        L.push("BEGIN:VALARM");
        L.push("TRIGGER:-PT" + remind + "M");
        L.push("ACTION:DISPLAY");
        L.push("DESCRIPTION:课程提醒");
        L.push("END:VALARM");
      }
      L.push("END:VEVENT");
    }
    L.push("END:VCALENDAR");
    return L.join("\r\n");
  }

  // ---------- 表格渲染 ----------
  var WEEKS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"];
  var currentFile = null;

  function weekName(w) { w = parseInt(w, 10); return (w >= 1 && w <= 7) ? WEEKS[w - 1] : "周一"; }
  function newRow(data) {
    d = data || {};
    var tr = document.createElement("tr");
    var cells = [
      '<td class="ctr"><input type="checkbox" class="del-check"></td>',
      '<td><input value="' + escAttr(d.name || "") + '"></td>',
      '<td><select>' + WEEKS.map(function (w) { return '<option' + (weekName(d.weekday) === w ? ' selected' : '') + '>' + w + '</option>'; }).join('') + '</select></td>',
      '<td><input value="' + escAttr(d.st || "08:00") + '"></td>',
      '<td><input value="' + escAttr(d.et || "09:35") + '"></td>',
      '<td><input value="' + escAttr(d.ws || "1") + '"></td>',
      '<td><input value="' + escAttr(d.we || "17") + '"></td>',
      '<td><input value="' + escAttr(d.loc || "") + '"></td>',
      '<td><input value="' + escAttr(d.teacher || "") + '"></td>',
      '<td><input value="' + escAttr(d.code || "") + '"></td>'
    ].join('');
    tr.innerHTML = cells;
    $("tbody").appendChild(tr);
  }
  function escAttr(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function clearRows() { $("tbody").innerHTML = ""; }
  function readRows() {
    var out = [];
    var rows = $("tbody").rows;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var inputs = r.querySelectorAll("input:not(.del-check), select");
      var name = inputs[0].value.trim();
      if (!name) continue;
      var weekday = WEEKS.indexOf(inputs[1].value) + 1 || 1;
      out.push({
        name: name,
        weekday: weekday,
        st: inputs[2].value.trim(),
        et: inputs[3].value.trim(),
        ws: parseInt(inputs[4].value, 10) || 1,
        we: parseInt(inputs[5].value, 10) || 17,
        loc: inputs[6].value.trim(),
        teacher: inputs[7].value.trim(),
        code: inputs[8].value.trim()
      });
    }
    return out;
  }

  // ---------- ICS 导入（Canvas 作业等） ----------
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  function parseIcsField(block, name) {
    var re = new RegExp("^" + name + "(?:;[^:\r\n]*)?:(.*)$", "m");
    var m = block.match(re);
    return m ? m[1].replace(/\r/g, "").trim() : "";
  }
  function parseIcsDt(s) {
    if (!s) return null;
    var m = s.match(/^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?(Z)?$/);
    if (!m) return null;
    var y = +m[1], mo = +m[2] - 1, d = +m[3], h = +(m[4] || 0), mi = +(m[5] || 0), se = +(m[6] || 0);
    if (m[7] === "Z") return new Date(Date.UTC(y, mo, d, h, mi, se));
    return new Date(y, mo, d, h, mi, se);
  }
  function parseIcs(text) {
    var out = [];
    var re = /BEGIN:VEVENT([\s\S]*?)END:VEVENT/g;
    var m;
    while ((m = re.exec(text))) {
      var b = m[1];
      var sd = parseIcsDt(parseIcsField(b, "DTSTART"));
      var ed = parseIcsDt(parseIcsField(b, "DTEND"));
      if (!sd) continue;
      if (!ed) ed = new Date(sd.getTime() + 30 * 60000);
      out.push({ name: parseIcsField(b, "SUMMARY") || "（未命名作业）", start: sd, end: ed });
    }
    return out;
  }
  function mapToRow(ev) {
    var fm = new Date($("firstMon").value + "T00:00:00");
    if (isNaN(fm.getTime())) return null;
    var local = new Date(ev.start.getTime());
    var loce = new Date(ev.end.getTime());
    var day0 = new Date(local.getFullYear(), local.getMonth(), local.getDate()).getTime();
    var fm0 = new Date(fm.getFullYear(), fm.getMonth(), fm.getDate()).getTime();
    var diff = Math.round((day0 - fm0) / 86400000);
    var ws = Math.floor(diff / 7) + 1;
    var wd = ((diff % 7) + 7) % 7 + 1;
    return {
      name: ev.name, weekday: wd, ws: ws, we: ws,
      st: pad2(local.getHours()) + ":" + pad2(local.getMinutes()),
      et: pad2(loce.getHours()) + ":" + pad2(loce.getMinutes()),
      loc: "", teacher: "", code: "Canvas"
    };
  }

  // ---------- 事件 ----------
  $("btnOpen").onclick = function () { $("fileInput").click(); };
  $("fileInput").onchange = async function (e) {
    var f = e.target.files[0];
    if (f) { currentFile = f; $("hint").textContent = "已选择：" + f.name + "（点「解析课表」开始）"; }
    e.target.value = "";
  };
  $("btnParse").onclick = async function () {
    if (!currentFile) { alert("请先选择一个 .docx 文件。"); return; }
    try {
      var buf = await currentFile.arrayBuffer();
      var xml = await extractDocumentXml(buf);
      var matrix = parseMatrix(xml);
      var entries = parseSchedule(matrix);
      clearRows();
      entries.forEach(newRow);
      $("hint").textContent = "解析完成：" + entries.length + " 个上课时段，请预览并修正后再导出。";
    } catch (err) { alert("解析失败：" + err.message); }
  };
  $("btnAdd").onclick = function () { newRow(); };
  $("btnDel").onclick = function () {
    var rows = $("tbody").rows;
    var n = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var cb = rows[i].querySelector(".del-check");
      if (cb && cb.checked) { $("tbody").removeChild(rows[i]); n++; }
    }
    $("hint").textContent = n ? ("已删除 " + n + " 行。") : "未选择要删除的行。";
  };
  $("checkAll").onchange = function () {
    var boxes = $("tbody").querySelectorAll(".del-check");
    for (var i = 0; i < boxes.length; i++) boxes[i].checked = this.checked;
  };
  $("btnExport").onclick = function () {
    var entries = readRows();
    if (!entries.length) { alert("没有可导出的课程。"); return; }
    var firstMon = new Date($("firstMon").value + "T00:00:00");
    if (isNaN(firstMon.getTime())) { alert("请选择正确的「第1周周一」日期。"); return; }
    var withTz = $("tz").value !== "floating";
    var remind = Math.max(0, parseInt($("remind").value, 10) || 0);
    var calName = $("calName").value.trim() || "课表";
    var ics = buildIcs(entries, firstMon, remind, calName, withTz);
    var blob = new Blob(["\uFEFF" + ics], { type: "text/calendar;charset=utf-8" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = calName + ".ics";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    $("hint").textContent = "已导出 " + entries.length + " 个时段到：" + calName + ".ics";
  };
  $("btnImportIcs").onclick = function () { $("icsInput").click(); };
  $("icsInput").onchange = async function (e) {
    var f = e.target.files[0];
    if (!f) return;
    try {
      var text = await f.text();
      var events = parseIcs(text);
      if (!events.length) { alert("该 .ics 里没有可导入的事件。"); e.target.value = ""; return; }
      var added = 0, skipped = 0;
      events.forEach(function (ev) {
        var row = mapToRow(ev);
        if (!row || row.ws < 1) { skipped++; return; }
        newRow(row); added++;
      });
      $("hint").textContent = "已导入 " + added + " 个作业事件" +
        (skipped ? ("（跳过 " + skipped + " 个早于第1周周一的事件，请先核对开学日期）") : "") +
        "，请核对日期后再导出。";
    } catch (err) { alert("导入失败：" + err.message); }
    e.target.value = "";
  };
  $("btnTutorial").onclick = function () { $("modal").showModal(); };
  $("modalClose").onclick = function () { $("modal").close(); };
})();