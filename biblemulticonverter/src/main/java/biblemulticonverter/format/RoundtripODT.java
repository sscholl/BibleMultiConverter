package biblemulticonverter.format;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import biblemulticonverter.data.Bible;
import biblemulticonverter.data.Book;
import biblemulticonverter.data.BookID;
import biblemulticonverter.data.Chapter;
import biblemulticonverter.data.FormattedText;
import biblemulticonverter.data.FormattedText.ExtraAttributePriority;
import biblemulticonverter.data.FormattedText.FormattingInstructionKind;
import biblemulticonverter.data.FormattedText.LineBreakKind;
import biblemulticonverter.data.FormattedText.RawHTMLMode;
import biblemulticonverter.data.FormattedText.Visitor;
import biblemulticonverter.data.Verse;

public class RoundtripODT implements RoundtripFormat {

	public static final String[] HELP_TEXT = {
			"ODT export and re-import [Unfinished]",
			"",
			"Usage: RoundtripODT <outfile>.odt [contrast|plain|printable|<styles.xml>]",
			"",
			"Export into an OpenDocument Text file. All features are exported, but some might look",
			"strange in OpenOffice/LibreOffice. The file can be edited in OpenOffice/LibreOffice and",
			"the resulting file can be converted back, without any loss of features in between.",
			"",
			"When editing the file, keep in mind that inline formatting is not parsed; in case you",
			"want to change the formatting, you have to use the existing paragraph and text",
			"styles, whose names start with BMC_.",
			"",
			"To verify a Bible, use the 'contrast' style; the 'plain' style is more friendly for",
			"reading; the 'printable' style hides all meta-information for printing. Or use your",
			"custom styles.xml."
	};

	private static final String OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
	private static final String TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
	private static final String XLINK = "http://www.w3.org/1999/xlink";

	@Override
	public void doExport(Bible bible, String... exportArgs) throws Exception {
		File exportFile = new File(exportArgs[0]);
		String styleName = exportArgs.length > 1 ? exportArgs[1] : "contrast";
		InputStream in = null;
		if (styleName.matches("[a-z]+")) {
			in = RoundtripODT.class.getResourceAsStream("/RoundtripODT/" + styleName + "_styles.xml");
		}
		if (in == null) {
			in = new FileInputStream(styleName);
		}
		try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(exportFile))) {
			ZipEntry mimetypeZE = new ZipEntry("mimetype");
			mimetypeZE.setSize(39);
			mimetypeZE.setCompressedSize(39);
			mimetypeZE.setCrc(204654174);
			mimetypeZE.setMethod(ZipOutputStream.STORED);
			zos.putNextEntry(mimetypeZE);
			zos.write("application/vnd.oasis.opendocument.text".getBytes(StandardCharsets.US_ASCII));
			zos.putNextEntry(new ZipEntry("content.xml"));
			int[] stats = new int[1];
			zos.write(toContentXml(bible, stats));
			zos.putNextEntry(new ZipEntry("styles.xml"));
			copyStream(in, zos);
			zos.putNextEntry(new ZipEntry("meta.xml"));
			try (InputStream metaIn = RoundtripODT.class.getResourceAsStream("/RoundtripODT/meta.xml")) {
				byte[] buf = new byte[1024];
				int len = metaIn.read(buf);
				int len2 = metaIn.read(buf, len, buf.length - len);
				if (len2 != -1)
					throw new IOException();
				String data = new String(buf, 0, len, StandardCharsets.US_ASCII).replace("#PARA#", "" + stats[0]);
				zos.write(data.getBytes(StandardCharsets.US_ASCII));
			}
			zos.putNextEntry(new ZipEntry("META-INF/manifest.xml"));
			copyStream(RoundtripODT.class.getResourceAsStream("/RoundtripODT/manifest.xml"), zos);
		}
	}

	private byte[] toContentXml(Bible bible, int[] stats) throws Exception {
		Map<String, Set<String>> xrefTargets = new HashMap<String, Set<String>>();
		XrefVisitor xrv = new XrefVisitor(xrefTargets);
		for (Book bk : bible.getBooks()) {
			for (Chapter ch : bk.getChapters()) {
				if (ch.getProlog() != null)
					ch.getProlog().accept(xrv);
				for (Verse v : ch.getVerses()) {
					v.accept(xrv);
				}
			}
		}
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		Document doc = dbf.newDocumentBuilder().newDocument();
		doc.appendChild(doc.createElementNS(OFFICE, "office:document-content"));
		doc.getDocumentElement().appendChild(doc.createElementNS(OFFICE, "office:body"));
		doc.getDocumentElement().setAttribute("xmlns:text", TEXT);
		doc.getDocumentElement().setAttribute("xmlns:xlink", XLINK);
		Element text = doc.createElementNS(OFFICE, "office:text");
		doc.getDocumentElement().getFirstChild().appendChild(text);
		RoundtripODTVisitor v = new RoundtripODTVisitor(text, false);
		Element p = doc.createElementNS(TEXT, "text:p");
		p.setAttributeNS(TEXT, "text:style-name", "BMC_5f_NextChapter");
		text.appendChild(p);
		appendSpan(p, "BMC_5f_Content", bible.getName());
		for (Book bb : bible.getBooks()) {
			p = doc.createElementNS(TEXT, "text:p");
			p.setAttributeNS(TEXT, "text:style-name", "BMC_5f_Book");
			text.appendChild(p);
			appendBookmarkTag(p, "start", "BMC-" + bb.getAbbr().replace('.', '_'));
			appendSpan(p, "BMC_5f_Verse", bb.getAbbr());
			appendBookmarkTag(p, "end", "BMC-" + bb.getAbbr().replace('.', '_'));
			appendSpan(p, "BMC_5f_Grammar", bb.getId().getOsisID());
			appendSpan(p, "BMC_5f_Ignored", " – ");
			if (!bb.getShortName().equals(bb.getLongName()))
				appendSpan(p, "BMC_5f_Special", bb.getShortName());
			appendSpan(p, "BMC_5f_Content", bb.getLongName());
			int cnumber = 0;
			for (Chapter ch : bb.getChapters()) {
				cnumber++;
				if (cnumber != 1) {
					p = doc.createElementNS(TEXT, "text:p");
					p.setAttributeNS(TEXT, "text:style-name", "BMC_5f_NextChapter");
					text.appendChild(p);
					appendSpan(p, "BMC_5f_Ignored", "– " + cnumber + " –");
				}
				if (ch.getProlog() != null) {
					ch.getProlog().accept(v);
					v.reset();
				}
				for (Verse vv : ch.getVerses()) {
					p = v.makeParagraph();
					String bmk = bb.getAbbr().replace('.', '_') + "-" + cnumber + "-" + vv.getNumber();
					Set<String> targets = xrefTargets.get(bmk);
					if (targets == null) {
						targets = Collections.emptySet();
					}
					appendBookmarkTag(p, "start", "BMC-" + bmk);
					for (String target : targets) {
						appendBookmarkTag(p, "start", "BMC-" + target);
					}
					appendSpan(p, "BMC_5f_Verse", vv.getNumber() + " ");
					appendBookmarkTag(p, "end", "BMC-" + bmk);
					for (String target : targets) {
						appendBookmarkTag(p, "end", "BMC-" + target);
					}
					vv.accept(v);
					v.reset();
				}
			}
		}
		stats[0] = doc.getElementsByTagNameNS(TEXT, "p").getLength();
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		TransformerFactory.newInstance().newTransformer().transform(new DOMSource(doc), new StreamResult(baos));
		return baos.toByteArray();
	}

	private static Element appendSpan(Element elem, String style, String value) {
		Element span = null;
		if (elem.getLastChild() != null && elem.getLastChild() instanceof Element) {
			Element last = (Element) elem.getLastChild();
			if (last.getLocalName().equals("span") && style.equals(last.getAttributeNS(TEXT, "style-name"))) {
				span = last;
			}
		}
		if (span == null) {
			span = elem.getOwnerDocument().createElementNS(TEXT, "text:span");
			span.setAttributeNS(TEXT, "text:style-name", style);
			elem.appendChild(span);
		}
		span.appendChild(elem.getOwnerDocument().createTextNode(value));
		return span;
	}

	private void appendBookmarkTag(Element elem, String type, String bookmark) {
		Element bmk = elem.getOwnerDocument().createElementNS(TEXT, "text:bookmark-" + type);
		bmk.setAttributeNS(TEXT, "text:name", bookmark);
		elem.appendChild(bmk);
	}

	private static Element appendLink(Element elem, String target) {
		Element link = elem.getOwnerDocument().createElementNS(TEXT, "text:a");
		link.setAttributeNS(XLINK, "xlink:type", "simple");
		link.setAttributeNS(XLINK, "xlink:href", target);
		elem.appendChild(link);
		return link;
	}

	private void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] buffer = new byte[4096];
		int len;
		while ((len = in.read(buffer)) != -1) {
			out.write(buffer, 0, len);
		}
		in.close();
	}

	@Override
	public Bible doImport(File inputDir) throws Exception {
		throw new UnsupportedOperationException("Import not yet implemented!");
	}

	@Override
	public boolean isExportImportRoundtrip() {
		return true;
	}

	@Override
	public boolean isImportExportRoundtrip() {
		return true;
	}

	public static class RoundtripODTVisitor implements FormattedText.Visitor<RuntimeException> {
		private final Element parent;
		private Element p;
		private String textStyle = "BMC_5f_Content", paragraphStyle;
		private final List<String> suffixStack = new ArrayList<String>();
		private FormattedText.FormattingInstructionKind pendingInstruction = null;

		public RoundtripODTVisitor(Element parent, boolean isParagraph) {
			this(parent, isParagraph, "BMC_5f_Content");
		}

		private RoundtripODTVisitor(Element parent, boolean isParagraph, String paragraphStyle) {
			this.parent = isParagraph ? null : parent;
			this.p = isParagraph ? parent : null;
			suffixStack.add(null);
			this.paragraphStyle = paragraphStyle;
		}

		public void reset() {
			if (suffixStack.size() != 0 || p != null || !paragraphStyle.equals("BMC_5f_Content") || !textStyle.equals("BMC_5f_Content") || pendingInstruction != null)
				throw new IllegalStateException();
			suffixStack.add(null);
		}

		@Override
		public int visitElementTypes(String elementTypes) throws RuntimeException {
			if (pendingInstruction == null)
				return 0;
			if (elementTypes == null)
				return 1;
			if (elementTypes.equals("t")) {
				switch (pendingInstruction) {
				case BOLD:
					textStyle = "BMC_5f_Content_2b_Bold";
					break;
				case ITALIC:
					textStyle = "BMC_5f_Content_2b_Italic";
					break;
				case DIVINE_NAME:
					textStyle = "BMC_5f_Content_2b_DivineName";
					break;
				case WORDS_OF_JESUS:
					textStyle = "BMC_5f_Content_2b_WOJ";
					break;
				default:
					throw new IllegalStateException(pendingInstruction.toString());
				}
				suffixStack.add("//");
			} else {
				appendSpan(p, "BMC_5f_Special", "<" + pendingInstruction.getCode() + ">");
				suffixStack.add("/");
			}
			pendingInstruction = null;
			return 0;
		}

		private Element makeParagraph() {
			if (p == null) {
				p = parent.getOwnerDocument().createElementNS(TEXT, "text:p");
				p.setAttributeNS(TEXT, "text:style-name", paragraphStyle);
				parent.appendChild(p);
			}
			return p;
		}

		@Override
		public Visitor<RuntimeException> visitHeadline(int depth) throws RuntimeException {
			p = null;
			paragraphStyle = "BMC_5f_Heading_5f_" + depth;
			makeParagraph();
			suffixStack.add(null);
			return this;
		}

		@Override
		public void visitStart() throws RuntimeException {
		}

		@Override
		public void visitText(String text) throws RuntimeException {
			appendSpan(makeParagraph(), textStyle, text);
		}

		@Override
		public Visitor<RuntimeException> visitFootnote() throws RuntimeException {
			makeParagraph();
			Element note = p.getOwnerDocument().createElementNS(TEXT, "text:note");
			appendSpan(p, textStyle, "").appendChild(note);
			note.setAttributeNS(TEXT, "text:note-class", "footnote");
			Element body = p.getOwnerDocument().createElementNS(TEXT, "text:note-body");
			note.appendChild(body);
			return new RoundtripODTVisitor(body, false, "Footnote");
		}

		@Override
		public Visitor<RuntimeException> visitCrossReference(String bookAbbr, BookID book, int firstChapter, String firstVerse, int lastChapter, String lastVerse) throws RuntimeException {
			return new RoundtripODTVisitor(appendLink(makeParagraph(), "#" + bookAbbr.replace('.', '_') + "-" + firstChapter + "-" + firstVerse + "-" + lastChapter + "-" + lastVerse), true);
		}

		@Override
		public Visitor<RuntimeException> visitFormattingInstruction(FormattingInstructionKind kind) throws RuntimeException {
			makeParagraph();
			switch (kind) {
			case BOLD:
			case ITALIC:
			case DIVINE_NAME:
			case WORDS_OF_JESUS:
				pendingInstruction = kind;
				break;
			case FOOTNOTE_LINK:
				return new RoundtripODTVisitor(appendLink(p, "#FootnoteLink"), true);
			case LINK:
				return new RoundtripODTVisitor(appendLink(p, "#Link"), true);
			default:
				appendSpan(p, "BMC_5f_Special", "<" + kind.getCode() + ">");
				suffixStack.add("/");
				break;
			}
			return this;
		}

		@Override
		public Visitor<RuntimeException> visitCSSFormatting(String css) throws RuntimeException {
			appendSpan(makeParagraph(), "BMC_5f_Special", "<css style=\"" + css + "\">");
			suffixStack.add("/");
			return this;
		}

		@Override
		public void visitVerseSeparator() throws RuntimeException {
			appendSpan(makeParagraph(), "BMC_5f_Special", "<vs>");
			visitText("/");
			appendSpan(makeParagraph(), "BMC_5f_Special", "/");
		}

		@Override
		public void visitLineBreak(LineBreakKind kind) throws RuntimeException {
			makeParagraph();
			switch (kind) {
			case NEWLINE:
				appendSpan(p, textStyle, "").appendChild(p.getOwnerDocument().createElementNS(TEXT, "text:line-break"));
				break;
			case NEWLINE_WITH_INDENT:
				Element span = appendSpan(p, textStyle, "");
				span.appendChild(p.getOwnerDocument().createElementNS(TEXT, "text:line-break"));
				span.appendChild(p.getOwnerDocument().createElementNS(TEXT, "text:tab"));
				break;
			case PARAGRAPH:
				p = null;
				makeParagraph();
				break;
			default:
				throw new IllegalArgumentException("Unsupported line break kind");
			}
		}

		@Override
		public Visitor<RuntimeException> visitGrammarInformation(int[] strongs, String[] rmac, int[] sourceIndices) throws RuntimeException {
			appendSpan(makeParagraph(), "BMC_5f_Grammar", "[");
			StringBuilder suffixBuilder = new StringBuilder("]");

			for (int i = 0; i < strongs.length; i++) {
				suffixBuilder.append(strongs[i]);
				if (rmac != null) {
					suffixBuilder.append(":" + rmac[i]);
					if (sourceIndices != null)
						suffixBuilder.append(":" + sourceIndices[i]);
				}
				suffixBuilder.append(' ');
			}
			suffixStack.add(suffixBuilder.toString());
			return this;
		}

		@Override
		public Visitor<RuntimeException> visitDictionaryEntry(String dictionary, String entry) throws RuntimeException {
			return new RoundtripODTVisitor(appendLink(makeParagraph(), dictionary + ".odt#BMC-" + entry.replace('.', '_')), true);
		}

		@Override
		public void visitRawHTML(RawHTMLMode mode, String raw) throws RuntimeException {
			int marker = 1;
			while (raw.contains("</raw:" + marker + ">")) {
				marker = (int) (Math.random() * 1000000);
			}
			appendSpan(makeParagraph(), "BMC_5f_Special", "<raw:" + marker + " mode=\"" + mode.name() + "\">" + raw + "</raw:" + marker + ">");
		}

		@Override
		public Visitor<RuntimeException> visitVariationText(String[] variations) throws RuntimeException {
			StringBuilder tag = new StringBuilder("<var vars=\"");
			for (int i = 0; i < variations.length; i++) {
				if (i > 0)
					tag.append(',');
				tag.append(variations[i]);
			}
			tag.append("\">");
			appendSpan(makeParagraph(), "BMC_5f_Special", tag.toString());
			suffixStack.add("/");
			return this;
		}

		@Override
		public Visitor<RuntimeException> visitExtraAttribute(ExtraAttributePriority prio, String category, String key, String value) throws RuntimeException {
			appendSpan(makeParagraph(), "BMC_5f_Special", "<extra prio=\"" + prio.name() + "\" category=\"" + category + "\" key=\"" + key + "\" value=\"" + value + "\">");
			suffixStack.add("/");
			return this;
		}

		@Override
		public boolean visitEnd() throws RuntimeException {
			String suffix = suffixStack.remove(suffixStack.size() - 1);
			if (suffix == null) {
				p = null;
				paragraphStyle = "BMC_5f_Content";
			} else if (suffix.equals("//")) {
				textStyle = "BMC_5f_Content";
			} else if (suffix.equals("/")) {
				appendSpan(p, "BMC_5f_Special", "/");
			} else {
				appendSpan(p, "BMC_5f_Grammar", suffix);
			}
			return false;
		}
	}

	public static class XrefVisitor extends FormattedText.VisitorAdapter<RuntimeException> {
		private final Map<String, Set<String>> xrefTargets;

		public XrefVisitor(Map<String, Set<String>> xrefTargets) {
			super(null);
			this.xrefTargets = xrefTargets;
		}

		@Override
		protected Visitor<RuntimeException> wrapChildVisitor(Visitor<RuntimeException> childVisitor) throws RuntimeException {
			return this;
		}

		@Override
		public Visitor<RuntimeException> visitCrossReference(String bookAbbr, BookID book, int firstChapter, String firstVerse, int lastChapter, String lastVerse) throws RuntimeException {
			String prefix = bookAbbr.replace('.', '_') + "-" + firstChapter + "-" + firstVerse;
			String full = prefix + "-" + lastChapter + "-" + lastVerse;
			Set<String> targets = xrefTargets.get(prefix);
			if (targets == null) {
				targets = new HashSet<String>();
				xrefTargets.put(prefix, targets);
			}
			targets.add(full);
			return this;
		}
	}
}
