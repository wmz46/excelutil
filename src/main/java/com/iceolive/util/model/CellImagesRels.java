package com.iceolive.util.model;

import com.iceolive.xpathmapper.annotation.XPath;
import lombok.Data;

import java.util.List;

/**
 * @author wangmianzhe
 */
@Data
public class CellImagesRels {
    @XPath("/Relationships/Relationship")
    private List<CellImageRels> cellImageRelsList;
    @Data
    public static class CellImageRels {
        @XPath("./@Id")
        private String rId;
        @XPath("./@Target")
        private String target;
    }
}
