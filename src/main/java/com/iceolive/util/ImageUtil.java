package com.iceolive.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

public class ImageUtil {
    public static BufferedImage Bytes2Image(byte[] bytes) {
        if(bytes == null){
            return null;
        }
        try {
            ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
            BufferedImage image = ImageIO.read(bais);
            bais.close();
            return image;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static byte[] Image2Bytes(BufferedImage image, String format) {
        if(image == null){
            return null;
        }
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(image, format, baos);
            byte[] bytes =  baos.toByteArray();
            baos.close();
            return bytes;

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
