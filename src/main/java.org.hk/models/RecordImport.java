package models;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import java.time.LocalDate;

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
    @Column(name = "date")
    private LocalDate date;
    @Column(name = "count")
    private double count;
    @Column(name = "sum")
    private double sum;
    @Column(name = "product")
    private String product;
    private Content content;
}

