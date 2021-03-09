package net.hanbd.luckyexcel4j.web;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * @author hanbd
 */
@Component
@ConfigurationProperties(prefix = "fs")
@Data
public class FsProperties {
    private String rootDir;
}
