//! Deploying the WPS add-in and starting the service with the user's session.

use std::path::{Path, PathBuf};
use wordgit_core::paths::{home_dir, AppPaths};
use wordgit_core::{Error, Result};

pub const ADDIN_NAME: &str = "WordGit";

/// WPS's per-user add-in folder.
pub fn jsaddons_dir() -> PathBuf {
    #[cfg(windows)]
    {
        let base = std::env::var_os("APPDATA").map(PathBuf::from).unwrap_or_else(|| home_dir().join("AppData/Roaming"));
        base.join("kingsoft").join("wps").join("jsaddons")
    }
    #[cfg(target_os = "macos")]
    {
        home_dir().join("Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons")
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        home_dir().join(".local/share/Kingsoft/wps/jsaddons")
    }
}

/// Default location of the add-in files shipped with the binary.
pub fn bundled_addin() -> Option<PathBuf> {
    let exe = std::env::current_exe().ok()?;
    let folder = exe.parent()?;
    [folder.join("addin"), folder.join("../share/wordgit/addin"), folder.join("../../addin-wps")]
        .into_iter()
        .find(|candidate| candidate.join("ribbon.xml").is_file())
}

fn copy_tree(source: &Path, target: &Path) -> std::io::Result<()> {
    std::fs::create_dir_all(target)?;
    for entry in std::fs::read_dir(source)? {
        let entry = entry?;
        let destination = target.join(entry.file_name());
        if entry.file_type()?.is_dir() {
            copy_tree(&entry.path(), &destination)?;
        } else {
            std::fs::copy(entry.path(), &destination)?;
        }
    }
    Ok(())
}

/// Copies the add-in into `jsaddons/WordGit_<version>/`, registers it in `publish.xml` the way
/// widely used community add-ins do, and records where the service file lives.
pub fn install_addin(source: &Path, paths: &AppPaths, version: &str) -> Result<PathBuf> {
    if !source.join("ribbon.xml").is_file() {
        return Err(Error::invalid(format!("找不到加载项文件：{}", source.display())));
    }
    let addons = jsaddons_dir();
    std::fs::create_dir_all(&addons)?;
    remove_installed_folders(&addons)?;
    let folder = addons.join(format!("{ADDIN_NAME}_{version}"));
    copy_tree(source, &folder)?;
    let config = serde_json::json!({
        "serviceFile": paths.root.join("service.json"),
        "home": paths.root,
        "version": version,
    });
    std::fs::write(folder.join("wordgit-config.json"), serde_json::to_vec_pretty(&config).unwrap_or_default())?;
    let record = format!(
        "<jsplugin url=\"http://127.0.0.1:{}/\" type=\"wps\" enable=\"enable_dev\" install=\"null\" version=\"{version}\" name=\"{ADDIN_NAME}\"/>",
        crate::service::DEFAULT_PORT
    );
    update_publish_xml(&addons.join("publish.xml"), Some(&record))?;
    Ok(folder)
}

pub fn uninstall_addin() -> Result<()> {
    let addons = jsaddons_dir();
    if addons.is_dir() {
        remove_installed_folders(&addons)?;
        update_publish_xml(&addons.join("publish.xml"), None)?;
    }
    Ok(())
}

fn remove_installed_folders(addons: &Path) -> Result<()> {
    for entry in std::fs::read_dir(addons)?.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        if entry.path().is_dir() && name.starts_with(&format!("{ADDIN_NAME}_")) {
            std::fs::remove_dir_all(entry.path())?;
        }
    }
    Ok(())
}

/// Replaces (or removes) our `<jsplugin ... name="WordGit"/>` record, keeping other add-ins.
pub fn update_publish_xml(path: &Path, record: Option<&str>) -> Result<()> {
    let existing = std::fs::read_to_string(path).unwrap_or_default();
    let mut lines: Vec<String> = Vec::new();
    let mut body = existing.as_str();
    let declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
    if let Some(start) = body.find("<jsplugins") {
        let after_open = body[start..].find('>').map(|i| start + i + 1).unwrap_or(body.len());
        let end = body.rfind("</jsplugins>").unwrap_or(body.len());
        body = if after_open <= end { &body[after_open..end] } else { "" };
        // Each record is a self-closing element; keep every record that is not ours.
        let mut rest = body;
        while let Some(open) = rest.find('<') {
            let Some(close) = rest[open..].find("/>") else { break };
            let element = &rest[open..open + close + 2];
            let ours = element.contains(&format!("name=\"{ADDIN_NAME}\""));
            if !ours && (element.starts_with("<jsplugin ") || element.starts_with("<jspluginonline ")) {
                lines.push(format!("  {element}"));
            }
            rest = &rest[open + close + 2..];
        }
    }
    if let Some(record) = record {
        lines.insert(0, format!("  {record}"));
    }
    let text = format!("{declaration}\n<jsplugins>\n{}{}</jsplugins>\n", lines.join("\n"), if lines.is_empty() { "" } else { "\n" });
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    std::fs::write(path, text)?;
    Ok(())
}

/// Registers the service to start with the user's session and returns a description.
pub fn enable_autostart(executable: &Path) -> Result<String> {
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        let autostart = home_dir().join(".config/autostart");
        std::fs::create_dir_all(&autostart)?;
        let entry = format!(
            "[Desktop Entry]\nType=Application\nName=WordGit 服务\nComment=WordGit 文档版本管理的本地服务\nExec=\"{}\" serve\nTerminal=false\nNoDisplay=true\nX-GNOME-Autostart-enabled=true\n",
            executable.display()
        );
        let file = autostart.join("wordgit.desktop");
        std::fs::write(&file, entry)?;
        Ok(format!("已添加登录自启动：{}", file.display()))
    }
    #[cfg(target_os = "macos")]
    {
        let agents = home_dir().join("Library/LaunchAgents");
        std::fs::create_dir_all(&agents)?;
        let file = agents.join("com.qth1225.wordgit.plist");
        let plist = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n<plist version=\"1.0\"><dict>\n<key>Label</key><string>com.qth1225.wordgit</string>\n<key>ProgramArguments</key><array><string>{}</string><string>serve</string></array>\n<key>RunAtLoad</key><true/>\n<key>KeepAlive</key><true/>\n</dict></plist>\n",
            executable.display()
        );
        std::fs::write(&file, plist)?;
        let _ = std::process::Command::new("launchctl").arg("load").arg("-w").arg(&file).status();
        Ok(format!("已添加登录自启动：{}", file.display()))
    }
    #[cfg(windows)]
    {
        let launcher = executable.with_file_name("wordgit-launcher.exe");
        if !launcher.is_file() {
            return Err(Error::invalid("缺少 WordGit 无窗口启动程序，请重新安装。"));
        }
        let value = format!("\"{}\"", launcher.display());
        let status = std::process::Command::new("reg")
            .args(["add", r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "/v", "WordGit", "/t", "REG_SZ", "/d"])
            .arg(&value)
            .arg("/f")
            .status()?;
        if !status.success() {
            return Err(Error::invalid("无法写入登录自启动注册表项。"));
        }
        Ok("已添加登录自启动（HKCU Run）。".into())
    }
}

pub fn disable_autostart() -> Result<()> {
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        wordgit_core::fsutil::try_remove(&home_dir().join(".config/autostart/wordgit.desktop"));
    }
    #[cfg(target_os = "macos")]
    {
        let file = home_dir().join("Library/LaunchAgents/com.qth1225.wordgit.plist");
        let _ = std::process::Command::new("launchctl").arg("unload").arg(&file).status();
        wordgit_core::fsutil::try_remove(&file);
    }
    #[cfg(windows)]
    {
        let _ = std::process::Command::new("reg")
            .args(["delete", r"HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "/v", "WordGit", "/f"])
            .status();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn publish_xml_keeps_other_addins() {
        let path = std::env::temp_dir().join(format!("wg-publish-{}.xml", wordgit_core::random_hex(4)));
        std::fs::write(
            &path,
            "<?xml version=\"1.0\"?>\n<jsplugins>\n<jsplugin url=\"http://127.0.0.1:3889/\" type=\"wps\" name=\"Zotero-Jsa\"/>\n<jsplugin url=\"old\" name=\"WordGit\"/>\n</jsplugins>",
        )
        .unwrap();
        update_publish_xml(&path, Some("<jsplugin url=\"new\" name=\"WordGit\"/>")).unwrap();
        let text = std::fs::read_to_string(&path).unwrap();
        assert!(text.contains("Zotero-Jsa"));
        assert!(text.contains("url=\"new\""));
        assert!(!text.contains("url=\"old\""));
        update_publish_xml(&path, None).unwrap();
        let text = std::fs::read_to_string(&path).unwrap();
        assert!(text.contains("Zotero-Jsa") && !text.contains("WordGit"));
        let _ = std::fs::remove_file(&path);
    }
}
