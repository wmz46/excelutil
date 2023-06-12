package com.iceolive.util.model;

import com.iceolive.xpathmapper.annotation.XPath;
import lombok.Data;

import java.util.List;

/**
 * @author wangmianzhe
 */
@Data
public class CellImages {
    @XPath("/etc:cellImages/etc:cellImage")
    private List<CellImage> cellImageList;

    @Data
    public static class CellImage {
        @XPath("./xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        private String id;
        @XPath("./xdr:pic/xdr:blipFill/a:blip/@embed")
        private String rId;
    }
}
