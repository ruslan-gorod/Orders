package org.hk.models;

import lombok.Data;

@Data
public class Raw {
    private String raw;
    private double count;

    public Raw(RecordImport recordImport) {
        this.raw = recordImport.getProduct();
        this.count = recordImport.getCount();
    }
}