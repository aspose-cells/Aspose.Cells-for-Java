package com.aspose.gridjsdemo.usermanagement.repository;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.aspose.gridjsdemo.usermanagement.entity.Role;
 
@Repository
public interface RoleRepository extends CrudRepository<Role, Long> {

    Role findByName(String role);
}