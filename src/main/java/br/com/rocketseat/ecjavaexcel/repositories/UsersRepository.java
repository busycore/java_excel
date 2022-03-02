package br.com.rocketseat.ecjavaexcel.repositories;

import br.com.rocketseat.ecjavaexcel.models.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;


@Repository
public interface UsersRepository extends JpaRepository<User,Integer> {
}
