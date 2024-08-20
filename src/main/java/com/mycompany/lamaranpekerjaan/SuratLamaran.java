/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.lamaranpekerjaan;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;

import java.io.FileOutputStream;
import java.util.Scanner;
import java.text.SimpleDateFormat;
import java.util.Date;
/**
 *
 * @author ASUS
 */
public class SuratLamaran {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Masukkan Nama: ");
        String nama = scanner.nextLine();
        
        System.out.print("Masukkan Alamat: ");
        String alamat = scanner.nextLine();
        
        System.out.print("Masukkan Nomor Telepon: ");
        String nomorTelepon = scanner.nextLine();
        
        System.out.print("Masukkan Emaail: ");
        String email = scanner.nextLine();

        System.out.print("Masukkan Tanggal: ");
        String tanggal = scanner.nextLine();

        System.out.print("Masukkan Nama Perusahaan: ");
        String namaPerusahaan = scanner.nextLine();

        System.out.print("Masukkan Alamat Perusahaan: ");
        String alamatPerusahaan = scanner.nextLine();

        System.out.print("Masukkan Tempat Tanggal Lahir: ");
        String ttl = scanner.nextLine();
        
        System.out.print("Masukkan Pendidikan: ");
        String pendidikan = scanner.nextLine();

        System.out.print("Masukkan Nama Posisi yang Dilamar: ");
        String posisi = scanner.nextLine();

        System.out.print("Masukkan Universitas: ");
        String universitas = scanner.nextLine();
        
        System.out.print("Masukkan Jurusan: ");
        String jurusan = scanner.nextLine();
        
        System.out.print("Masukkan Ipk: ");
        String ipk = scanner.nextLine();

        System.out.print("Masukkan Bidang: ");
        String bidang = scanner.nextLine();

        System.out.print("Masukkan Keterampilan yang Relevan: ");
        String keterampilan = scanner.nextLine();
        
        // Menambahkan timestamp ke nama file
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        
        // Meminta nama file untuk disimpan
        System.out.print("Masukkan nama file (tanpa ekstensi): ");
        String baseFileName = scanner.nextLine();
        String wordFileName = baseFileName + "_" + timestamp + ".docx"; // Tambahkan timestamp
        String pdfFileName = baseFileName + "_" + timestamp + ".pdf"; // Tambahkan timestamp

            // Membuat isi surat lamaran
            String suratLamaran = String.format(
                "%s\n%s\n%s\n%s\n%s\n\nKepada Yth.\nHRD %s\n%s\n\nDengan hormat,\n\n" +
                "Saya yang bertanda tangan di bawah ini:\n\n" +
                "Nama: %s\n" +
                "Tempat/Tanggal Lahir: %s\n" +
                "Pendidikan Terakhir: %s\n" +
                "Alamat: %s\n\n" +
                "Dengan ini mengajukan lamaran kerja untuk posisi %s di perusahaan %s. Saya adalah lulusan %s jurusan %s " +
                "dengan IPK %s. Selama masa studi, saya telah mengembangkan kemampuan dan pengetahuan di bidang %s " +
                "melalui berbagai proyek, magang, dan kegiatan organisasi.\n\n" +
                "Saya juga memiliki keterampilan %s, yang saya yakin dapat memberikan kontribusi positif bagi %s. " +
                "Saya sangat termotivasi untuk belajar dan siap bekerja keras guna mencapai target yang telah ditetapkan " +
                "oleh perusahaan.\n\n" +
                "Sebagai bahan pertimbangan, berikut saya lampirkan beberapa dokumen pendukung:\n\n" +
                "   1. Curriculum Vitae (CV)\n" +
                "   2. Fotokopi Ijazah dan Transkrip Nilai\n" +
                "\nSaya berharap dapat diberikan kesempatan untuk wawancara agar dapat menjelaskan lebih rinci tentang " +
                "potensi yang saya miliki. Saya sangat antusias untuk menjadi bagian dari %s dan siap memberikan kontribusi terbaik saya.\n\n" +
                "Demikian surat lamaran ini saya sampaikan. Atas perhatian dan kesempatan yang diberikan, saya ucapkan terima kasih.\n\n" +
                "\n\nHormat saya,\n\n\n%s",
                nama, alamat, nomorTelepon, email, tanggal, namaPerusahaan, alamatPerusahaan, nama, ttl,
                pendidikan, alamat, posisi, namaPerusahaan, universitas, jurusan, ipk, bidang,
                keterampilan, namaPerusahaan, namaPerusahaan, nama
            );
            // Membuat file Word
            createWordFile(suratLamaran, wordFileName);

            // Membuat file PDF
            createPdfFile(suratLamaran, pdfFileName);
    }

    // Method untuk membuat file Word
    public static void createWordFile(String suratLamaran, String fileName) {
        try (XWPFDocument document = new XWPFDocument()) {
            // Membuat paragraf dalam dokumen Word
            String[] lines = suratLamaran.split("\n");

            for (String line : lines) {
                // Membuat paragraf baru untuk setiap baris
                XWPFParagraph paragraph = document.createParagraph();
                paragraph.setAlignment(org.apache.poi.xwpf.usermodel.ParagraphAlignment.BOTH); // Mengatur justify
                XWPFRun run = paragraph.createRun();
                paragraph.setSpacingAfter(0);
                run.setText(line);
                run.setFontSize(12); // Set font size
                run.setFontFamily("Times New Roman");
            }
             // Set font family

            // Menyimpan dokumen Word ke file
            try (FileOutputStream out = new FileOutputStream(fileName)) {
                document.write(out);
                System.out.println("Surat Lamaran berhasil dibuat sebagai file Word.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Method untuk membuat file PDF
    public static void createPdfFile(String suratLamaran, String fileName) {
        try {
            // Membuat dokumen PDF
            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(fileName));
            document.open();
            
            Font font = FontFactory.getFont(FontFactory.TIMES_ROMAN, 12, Font.NORMAL);
            
            // Mengatur spacing
            float leading = 14f; // Spacing antar baris
            float spacingBefore = 2f; // Spacing sebelum paragraf
            float spacingAfter = 2f; // Spacing setelah paragraf
            
            Paragraph paragraph = new Paragraph(suratLamaran, font);
            paragraph.setLeading(leading);
            paragraph.setSpacingBefore(spacingBefore);
            paragraph.setSpacingAfter(spacingAfter);
            paragraph.setAlignment(Paragraph.ALIGN_JUSTIFIED); // Mengatur justify

            // Menambahkan paragraf ke PDF
            document.add(paragraph);

            // Menutup dokumen
            document.close();
            System.out.println("Surat Lamaran berhasil dibuat sebagai file PDF.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
