package net.hanbd.luckyexcel4j.web.controller;

import cn.hutool.core.util.IdUtil;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import net.hanbd.luckyexcel4j.lucky.poi.LuckyFile;
import net.hanbd.luckyexcel4j.lucky.poi.base.FileMeta;
import net.hanbd.luckyexcel4j.web.FsProperties;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Nullable;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;

/**
 * @author hanbd
 */
@RestController
@Slf4j
public class IndexController {

    private final Path rootPath;

    @Autowired
    public IndexController(FsProperties fsProperties) {
        rootPath = Paths.get(fsProperties.getRootDir()).toAbsolutePath().normalize();
        try {
            Files.createDirectories(this.rootPath);
        } catch (Exception e) {
            log.error("create dir failed,Path:{}", this.rootPath);
            throw new RuntimeException("create dir failed,Path:" + this.rootPath, e);
        }
    }

    @SneakyThrows
    @GetMapping("/file/meta")
    public FileMeta getDemoFileMeta(@RequestParam(value = "filename", required = false) @Nullable String filename) {
        if (StringUtils.isEmpty(filename)) {
            filename = "index.xlsx";
        }
        Path filePath = rootPath.resolve(filename);
        if (!filePath.toFile().exists()) {
            throw new RuntimeException("对应文件不存在");
        }

        LuckyFile luckyFile = new LuckyFile(filePath.toFile(), PackageAccess.READ);
        return luckyFile.parse();
    }

    @SneakyThrows
    @PostMapping(value = "/file", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public String uploadFile(@RequestPart("excel") MultipartFile file) {
        String originalFilename = file.getOriginalFilename();
        String ext = StringUtils.getFilenameExtension(originalFilename);
        if (!StringUtils.hasText(ext) || !Objects.equals(ext.trim().toLowerCase(), "xlsx")) {
            throw new IllegalArgumentException("只支持xlsx文件上传");
        }

        File dest = rootPath.resolve(IdUtil.fastSimpleUUID() + "." + ext).toFile();
        file.transferTo(dest);

        return dest.getName();
    }
}
