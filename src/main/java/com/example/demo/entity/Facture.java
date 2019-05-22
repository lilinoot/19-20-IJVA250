package com.example.demo.entity;

import javax.persistence.*;
import java.util.Set;

@Entity
public class Facture {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @ManyToOne // Pour dire à Spring que ma facture fait référence à un client, il crée automatiquement une colonne qui est une clé étrangère de client
    private Client client;

    @OneToMany(mappedBy = "facture") // Fait référence à plusieurs factures
    private Set<LigneFacture> ligneFactures; // Set = une liste sans doublon, une collection

    public Long getId() { return id;}

    public void setId(Long id) {
        this.id = id;
    }

    public Client getClient() { return client; }

    public void setClient(Client client) { this.client = client; }

    public Set<LigneFacture> getLigneFactures() { return ligneFactures; }

    public void setLigneFactures(Set<LigneFacture> ligneFactures) { this.ligneFactures = ligneFactures; }

    public Double getTotal() {
        Double total = 0.0;
        for (LigneFacture ligneFacture : ligneFactures) {
            total = total + ligneFacture.getSousTotal();
        }
        return total;
    }
}

