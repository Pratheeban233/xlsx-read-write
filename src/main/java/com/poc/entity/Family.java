package com.poc.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.*;

@Entity
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Family {

	@Id
	@Column(name = "Sl.No",unique = true)
	private Long id;

	private String name;

	private String relation;
}
