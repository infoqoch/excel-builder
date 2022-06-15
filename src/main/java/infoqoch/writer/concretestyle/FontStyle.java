package infoqoch.writer.concretestyle;

import org.apache.poi.ss.usermodel.Font;

public enum FontStyle {
	NONE, BOLD, ITALIC;

	void applyFont(Font font) {
		if(this.equals(NONE))
			return;
		if(this.equals(BOLD))
			font.setBold(true);
		if(this.equals(ITALIC))
			font.setItalic(true);
	}
}