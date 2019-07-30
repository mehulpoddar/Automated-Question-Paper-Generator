import org.apache.poi.xwpf.usermodel.XWPFPictureData;

public class PaperElement {
	private int EleID;
	private String Content;
	private XWPFPictureData Img;
	
	PaperElement(int id, String c)
	{
		EleID = id;
		Content = c;
	}
	PaperElement(int id, XWPFPictureData img)
	{
		EleID = id;
		Img = img;
	}
	int getID()
	{
		return EleID;
	}
	String getContent()
	{
		return Content;
	}
	XWPFPictureData getImage()
	{
		return Img;
	}
}
