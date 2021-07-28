<%@ Page Language="C#" EnableViewState="false" EnableSessionState="false" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.Drawing.Drawing2D" %>
<script runat="server">
// Smart Image Processor
// Version: 1.1.5

void Page_Load(Object s, EventArgs e) {
    int intNewWidth, intNewHeight, maxWidth = 10000, maxHeight = 10000, qQuality = 80;
    if (Request["w"] != null) maxWidth = int.Parse(Request["w"]);
    if (Request["h"] != null) maxHeight = int.Parse(Request["h"]);
    if (Request["q"] != null) qQuality = int.Parse(Request["q"]);
    
    //get image from parameter
    string pictureFileName = Request["f"];
    string newFileName = Request["nf"];
    if (pictureFileName == null || pictureFileName == "" || !System.IO.File.Exists(pictureFileName)) {
        Response.Write("Error: File (" + pictureFileName + ") not found or empty");  
        return;
    }
    System.Drawing.Image inputImage = System.Drawing.Image.FromFile(pictureFileName);
    
    //define size for new image
    string aspect = Request["a"];
    if (aspect == "true") {
        if (maxWidth < inputImage.Width || maxHeight < inputImage.Height) {
            if (maxWidth >= maxHeight) {
                intNewWidth = (int)((double)maxHeight*((double)inputImage.Width/(double)inputImage.Height));
                intNewHeight = maxHeight;
            } else {
                intNewWidth = maxWidth;
                intNewHeight = (int)((double)maxWidth*((double)inputImage.Height/(double)inputImage.Width));
            }
            if (intNewWidth > maxWidth) {
                intNewWidth = maxWidth;
                intNewHeight = (int)((double)maxWidth*((double)inputImage.Height/(double)inputImage.Width));
            }
            if (intNewHeight > maxHeight) {
                intNewWidth = (int)((double)maxHeight*((double)inputImage.Width/(double)inputImage.Height));
                intNewHeight = maxHeight;
            }
        } else {
            intNewWidth = inputImage.Width;
            intNewHeight = inputImage.Height;
        }
    } else {
            intNewWidth = maxWidth;
            intNewHeight = maxHeight;
    }
    
    try {        
        //output new image with different size
        System.Drawing.Bitmap outputBitMap = new System.Drawing.Bitmap(intNewWidth, intNewHeight);
        System.Drawing.Graphics imageGraph = System.Drawing.Graphics.FromImage(outputBitMap);
        //imageGraph.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        imageGraph.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
        imageGraph.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        imageGraph.DrawImage(inputImage, -1, -1, outputBitMap.Width+1, outputBitMap.Height+1);
        inputImage.Dispose();
        System.Drawing.Imaging.EncoderParameters eps = new System.Drawing.Imaging.EncoderParameters(1);
        eps.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qQuality);
        System.Drawing.Imaging.ImageCodecInfo ici = GetEncoderInfo("image/jpeg");
        if (pictureFileName.ToLower() == newFileName.ToLower())
            System.IO.File.Delete(pictureFileName);
        outputBitMap.Save(newFileName, ici, eps);
        outputBitMap.Dispose();
    } catch (Exception ex) {
        Response.Write("Error: " + ex);
        return;
    }
    
    Response.Write(intNewWidth + ";" + intNewHeight + ";" + "DONE");
}

private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(String mimeType)
    {
    int j;
    System.Drawing.Imaging.ImageCodecInfo[] encoders;
    encoders = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
    for(j = 0; j < encoders.Length; ++j) {
        if(encoders[j].MimeType == mimeType)
            return encoders[j];
    }
    return null;
}
    
</script>