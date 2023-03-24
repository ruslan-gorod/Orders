package org.hk.models;

import lombok.Data;
import org.hibernate.Session;

@Data
public class QueryParameters {
    private Session session;
    private String dt;
    private String kt;
    private String document;
}