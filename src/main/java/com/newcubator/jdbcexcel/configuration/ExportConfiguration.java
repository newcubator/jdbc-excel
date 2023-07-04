package com.newcubator.jdbcexcel.configuration;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.RequiredArgsConstructor;

@AllArgsConstructor
@RequiredArgsConstructor
@Data
public class ExportConfiguration {

    private boolean autogenerateHyperlinks = true;
}
