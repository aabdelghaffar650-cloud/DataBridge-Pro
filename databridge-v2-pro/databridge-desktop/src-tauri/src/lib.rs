use tauri_plugin_shell::ShellExt;

  #[cfg_attr(mobile, tauri::mobile_entry_point)]
  pub fn run() {
      tauri::Builder::default()
          .plugin(tauri_plugin_opener::init())
          .plugin(tauri_plugin_shell::init())
          .setup(|app| {
              let sidecar = app.shell().sidecar("databridge-backend")
                  .expect("failed to create sidecar command");
              let (mut _rx, _child) = sidecar.spawn()
                  .expect("failed to spawn databridge-backend sidecar");
              Ok(())
          })
          .run(tauri::generate_context!())
          .expect("error while running tauri application");
  }