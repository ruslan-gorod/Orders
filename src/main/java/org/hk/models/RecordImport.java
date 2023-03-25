package org.hk.models;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Transient;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Data
@Entity
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class RecordImport {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "id", nullable = false)
    private Long id;
    @Column(name = "dt")
    private String dt;
    @Column(name = "kt")
    private String kt;
    @Column(name = "originDocument")
    private String originDocument;
    @Column(name = "compareDocument")
    private String compareDocument;
    @Column(name = "criteriaDocument")
    private String criteriaDocument;
    @Column(name = "date")
    private LocalDate date;
    @Column(name = "count")
    private double count;
    @Column(name = "countResult")
    private double countResult;
    @Column(name = "sum")
    private double sum;
    @Column(name = "product")
    private String product;
    @Column(name = "partner")
    private String partner;
    @Transient
    private Content content;
    @Transient
    private List<Raw> rawList = new ArrayList<>();
}