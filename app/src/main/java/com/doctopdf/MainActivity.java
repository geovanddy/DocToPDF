package com.doctopdf;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.View;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import java.io.File;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends AppCompatActivity {

    private TextView tvStatus;
    private TextView tvFileName;
    private ProgressBar progressBar;
    private Button btnSelectFile;
    private Button btnOpenDrive;
    private ImageView ivFileIcon;

    private File convertedPdf = null;
    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final Handler mainHandler = new Handler(Looper.getMainLooper());

    private final ActivityResultLauncher<String[]> filePickerLauncher =
        registerForActivityResult(new ActivityResultContracts.OpenDocument(), uri -> {
            if (uri != null) {
                handleFileUri(uri);
            }
        });

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        tvStatus = findViewById(R.id.tv_status);
        tvFileName = findViewById(R.id.tv_filename);
        progressBar = findViewById(R.id.progress_bar);
        btnSelectFile = findViewById(R.id.btn_select_file);
        btnOpenDrive = findViewById(R.id.btn_open_drive);
        ivFileIcon = findViewById(R.id.iv_file_icon);

        btnSelectFile.setOnClickListener(v -> {
            filePickerLauncher.launch(new String[]{
                "application/msword",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/vnd.ms-excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-powerpoint",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            });
        });

        btnOpenDrive.setOnClickListener(v -> {
            if (convertedPdf != null && convertedPdf.exists()) {
                openInGoogleDrive(convertedPdf);
            }
        });

        // Manejar si la app se abre con un archivo compartido
        handleIncomingIntent(getIntent());
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        handleIncomingIntent(intent);
    }

    private void handleIncomingIntent(Intent intent) {
        if (intent == null) return;
        String action = intent.getAction();
        Uri data = intent.getData();

        if (data != null && (Intent.ACTION_VIEW.equals(action) || Intent.ACTION_SEND.equals(action))) {
            handleFileUri(data);
        } else if (Intent.ACTION_SEND.equals(action)) {
            Uri uri = intent.getParcelableExtra(Intent.EXTRA_STREAM);
            if (uri != null) handleFileUri(uri);
        }
    }

    private void handleFileUri(Uri uri) {
        String fileName = FileUtils.getFileName(this, uri);
        if (fileName == null) fileName = "archivo";

        final String finalFileName = fileName;
        setStatus("⏳ Convirtiendo: " + fileName, true);
        tvFileName.setText(finalFileName);
        setFileIcon(fileName);
        btnOpenDrive.setEnabled(false);
        convertedPdf = null;

        executor.execute(() -> {
            try {
                File outputPdf = ConversionUtils.convertToPdf(MainActivity.this, uri, finalFileName);
                mainHandler.post(() -> {
                    convertedPdf = outputPdf;
                    setStatus("✅ ¡Conversión exitosa!", false);
                    tvFileName.setText(outputPdf.getName());
                    btnOpenDrive.setEnabled(true);
                    openInGoogleDrive(outputPdf);
                });
            } catch (Exception e) {
                mainHandler.post(() -> {
                    setStatus("❌ Error: " + e.getMessage(), false);
                    Toast.makeText(MainActivity.this, "Error al convertir el archivo", Toast.LENGTH_LONG).show();
                });
            }
        });
    }

    private void openInGoogleDrive(File pdfFile) {
        try {
            Uri pdfUri = FileProvider.getUriForFile(this,
                getPackageName() + ".fileprovider", pdfFile);

            Intent intent = new Intent(Intent.ACTION_VIEW);
            intent.setDataAndType(pdfUri, "application/pdf");
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            intent.addFlags(Intent.FLAG_ACTIVITY_NO_HISTORY);

            // Intentar abrir con Google Drive específicamente
            intent.setPackage("com.google.android.apps.docs");

            try {
                startActivity(intent);
            } catch (Exception e) {
                // Si Drive no está, abrir con cualquier visor de PDF
                intent.setPackage(null);
                startActivity(Intent.createChooser(intent, "Abrir PDF con..."));
            }
        } catch (Exception e) {
            Toast.makeText(this, "Error al abrir el PDF: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void setStatus(String message, boolean loading) {
        tvStatus.setText(message);
        progressBar.setVisibility(loading ? View.VISIBLE : View.GONE);
    }

    private void setFileIcon(String fileName) {
        String lower = fileName.toLowerCase();
        if (lower.endsWith(".doc") || lower.endsWith(".docx")) {
            ivFileIcon.setImageResource(R.drawable.ic_word);
        } else if (lower.endsWith(".xls") || lower.endsWith(".xlsx")) {
            ivFileIcon.setImageResource(R.drawable.ic_excel);
        } else if (lower.endsWith(".ppt") || lower.endsWith(".pptx")) {
            ivFileIcon.setImageResource(R.drawable.ic_powerpoint);
        } else {
            ivFileIcon.setImageResource(R.drawable.ic_file);
        }
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        executor.shutdown();
    }
}
