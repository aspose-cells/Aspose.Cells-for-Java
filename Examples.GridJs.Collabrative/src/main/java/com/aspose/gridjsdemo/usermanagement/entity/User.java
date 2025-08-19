package com.aspose.gridjsdemo.usermanagement.entity;

import java.io.Serializable;
import java.time.LocalDateTime;
import java.util.Set;

import org.hibernate.annotations.GenericGenerator;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.annotation.LastModifiedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

//import jakarta.persistence.Column;
//import jakarta.persistence.Entity;
//import jakarta.persistence.EntityListeners;
//import jakarta.persistence.FetchType;
//import jakarta.persistence.GeneratedValue;
//import jakarta.persistence.GenerationType;
//import jakarta.persistence.Id;
//import jakarta.persistence.JoinColumn;
//import jakarta.persistence.JoinTable;
//import jakarta.persistence.ManyToMany;
//import jakarta.persistence.Transient;
//import jakarta.validation.constraints.NotBlank;
//import jakarta.validation.constraints.Size;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.EntityListeners;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.JoinTable;
import javax.persistence.ManyToMany;
import javax.persistence.Transient;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.Size;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
 
@Entity
@Data
@AllArgsConstructor
@NoArgsConstructor
@EqualsAndHashCode
@EntityListeners(AuditingEntityListener.class)
public class User implements Serializable {

   
 
	private static final long serialVersionUID = 1L;

	@Id
    @GeneratedValue(strategy = GenerationType.AUTO, generator = "native")
    @GenericGenerator(name = "native", strategy = "native")
    private Long id;

    @Column
    @NotBlank
    @Size(min = 2, max = 15, message = "Size rules are not followed")
    private String firstName;

    @Column
    @NotBlank
    @Size(min = 2, max = 15, message = "Size rules are not followed")
    private String lastName;

    @Column
    @NotBlank
    private String email;

    @Column
    @NotBlank
    private String username;

    @Column
    @NotBlank
    private String password;

    @Transient
    private String confirmPassword;

    @Size(min = 1)
    @ManyToMany(fetch = FetchType.LAZY)
    @JoinTable(name = "user_roles",
            joinColumns = @JoinColumn(name = "user_id"),
            inverseJoinColumns = @JoinColumn(name = "role_id"))
    private Set<Role> roles;

    @CreatedDate
    private LocalDateTime createdOn;

    @LastModifiedDate
    private LocalDateTime lastModifiedOn;
}
