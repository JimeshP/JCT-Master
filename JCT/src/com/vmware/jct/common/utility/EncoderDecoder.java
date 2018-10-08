package com.vmware.jct.common.utility;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;
/**
 * Class is used for encoding and decoding
 * @author InterraIT
 *
 */
public class EncoderDecoder {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(EncoderDecoder.class);
	/**
     * Decode string to image
     * @param imageString The string to decode
     * @return decoded image
     */
    public static byte[] getImageBytes(String imageString) {
    	LOGGER.info(">>>> EncoderDecoder.getImageBytes");
        byte[] imageByte = null;
        try {
            BASE64Decoder decoder = new BASE64Decoder();
            imageByte = decoder.decodeBuffer(imageString);
            ByteArrayInputStream bis = new ByteArrayInputStream(imageByte);
            ImageIO.read(bis);
            bis.close();
        } catch (Exception e) {
           // e.printStackTrace();
        	LOGGER.error(e.getLocalizedMessage());
        }
        LOGGER.info("<<<< EncoderDecoder.getImageBytes");
        return imageByte;
    }
    /**
     * Reads the byte [] and returns base 64 string 
     * @param imageByte
     * @return base64 representation of image
     */
    public static String getBase64String(byte[] imageByte){
    	LOGGER.info(">>>> EncoderDecoder.getBase64String");
    	BufferedImage image = getBufferedImage(imageByte);
    	String str = encodeToString(image, "png");
    	LOGGER.info("<<<< EncoderDecoder.getBase64String");
    	return str;
    }
    private static BufferedImage getBufferedImage(byte[] imageByte) {
    	LOGGER.info(">>>> EncoderDecoder.getBufferedImage");
        BufferedImage image = null;
        try {
            ByteArrayInputStream bis = new ByteArrayInputStream(imageByte);
            image = ImageIO.read(bis);
            bis.close();
        } catch (Exception e) {
           // e.printStackTrace();
        	LOGGER.error(e.getLocalizedMessage());
        }
        LOGGER.info("<<<< EncoderDecoder.getBufferedImage");
        return image;
    }

    /**
     * Encode image to string
     * @param image The image to encode
     * @param type jpeg, bmp, ...
     * @return encoded string
     */
    public static String encodeToString(BufferedImage image, String type) {
    	LOGGER.info(">>>> EncoderDecoder.encodeToString");
        String imageString = null;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        try {
            ImageIO.write(image, type, bos);
            byte[] imageBytes = bos.toByteArray();

            BASE64Encoder encoder = new BASE64Encoder();
            imageString = encoder.encode(imageBytes);

            bos.close();
        } catch (IOException e) {
           // e.printStackTrace();
        	LOGGER.error(e.getLocalizedMessage());
        }
        LOGGER.info("<<<< EncoderDecoder.encodeToString");
        return imageString;
    }
}
