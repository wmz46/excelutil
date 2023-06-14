package com.iceolive.util.exception;

public class ImageOutOfBoundsException extends RuntimeException{
    public ImageOutOfBoundsException(){
        super("图片越界");
    }
}
