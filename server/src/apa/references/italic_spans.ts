export interface ItalicSpan { start: number; end: number }

/** Restore only recognizable APA reference emphasis; all offsets refer to
 * the unchanged input. Unknown containers are left for manual formatting. */
export function referenceItalicSpans(text: string) {
  const spans: ItalicSpan[] = [];
  const warnings: string[] = [];
  const heading = /^(?:references|bibliography|works cited)\s*:?\s*$/im.exec(text);
  const minimum = heading ? heading.index + heading[0].length : 0;
  const starts = [...text.matchAll(/^[^\r\n]+?\(((?:1[6-9]|20)\d{2}[a-z]?|n\.d\.)\)\.\s+/gm)]
    .filter((m) => m.index! >= minimum && (/\b[A-Z]\.\s*/.test(m[0]) || /\b(?:Institute|University|Department|Association|Organization|Bureau|Council|Fund|Agency|Commission|Foundation)\b/i.test(m[0])));
  for (let i = 0; i < starts.length; i++) {
    const match = starts[i]!;
    const start = match.index! + match[0].length;
    const end = starts[i + 1]?.index ?? text.length;
    let depth = 0;
    let titleEnd = -1;
    for (let j = start; j < end; j++) {
      if (text[j] === "(") depth++;
      if (text[j] === ")") depth = Math.max(0, depth - 1);
      if (depth === 0 && /[.!?]/.test(text[j]!) && (j + 1 === end || /\s/.test(text[j + 1]!))) { titleEnd = j; break; }
    }
    if (titleEnd < start) continue;
    const tailStart = titleEnd + 1;
    const tail = text.slice(tailStart, end);
    const journal = /^\s*([^,.!?]+),\s*(\d+)\s*(?=\(|,|\.)/.exec(tail);
    const onlineJournal = /^\s*([^.!?]+)[.!?]\s*(?=https?:\/\/|doi:|$)/i.exec(tail);
    if (journal) {
      const name = journal[1]!.trim();
      const journalStart = tailStart + journal[0].indexOf(name);
      const volumeStart = tailStart + journal[0].lastIndexOf(journal[2]!);
      spans.push({ start: journalStart, end: journalStart + name.length }, { start: volumeStart, end: volumeStart + journal[2]!.length });
    } else if (onlineJournal && /\b(?:Journal|Review|Annals|Proceedings|Transactions|Quarterly|Bulletin)\b/i.test(onlineJournal[1]!)) {
      const name = onlineJournal[1]!.trim();
      const index = tailStart + onlineJournal[0].indexOf(name);
      spans.push({ start: index, end: index + name.length });
    } else if (/\b(?:report|framework|fact sheet|guideline|manual|handbook)\b/i.test(text.slice(start, titleEnd)) ||
      /\b(?:Press|Books|Publishing|Publishers|Springer|Routledge|Wiley|Pearson)\b/.test(tail)) {
      if (!/\bIn\s+[^\n]*\(Eds?\./.test(tail)) spans.push({ start, end: titleEnd });
    } else warnings.push(`Check italics for the reference beginning “${text.slice(match.index!, match.index! + 65)}…”. Its source type is unclear.`);
  }
  if (!starts.length) warnings.push("No recognizable reference entries were found. Include authors, a year in parentheses, and the title, or use the Italic button.");
  return { spans, warnings };
}
