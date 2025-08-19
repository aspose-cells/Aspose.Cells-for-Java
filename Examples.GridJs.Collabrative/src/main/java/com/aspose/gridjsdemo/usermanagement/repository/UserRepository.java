package com.aspose.gridjsdemo.usermanagement.repository;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.aspose.gridjsdemo.usermanagement.entity.User;

import java.util.Optional;
 
@Repository
public interface UserRepository extends CrudRepository<User, Long> {

    Optional<User> findByUsername(String username);
}