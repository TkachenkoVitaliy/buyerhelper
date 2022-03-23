package com.tkachenko.buyerhelper.repository;

import com.tkachenko.buyerhelper.entity.FileEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface FileEntityRepository extends JpaRepository <FileEntity, Long> {

    FileEntity findByStorageFileNameAndIsActual(String storageFileName, boolean isTrue);
}
