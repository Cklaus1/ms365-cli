import { describe, it, expect } from 'vitest';
import { detectMimeType } from '../src/mime.js';

describe('detectMimeType', () => {
  it('detects common document types', () => {
    expect(detectMimeType('report.pdf')).toBe('application/pdf');
    expect(detectMimeType('doc.docx')).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    expect(detectMimeType('sheet.xlsx')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(detectMimeType('slides.pptx')).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  });

  it('detects image types', () => {
    expect(detectMimeType('photo.jpg')).toBe('image/jpeg');
    expect(detectMimeType('photo.jpeg')).toBe('image/jpeg');
    expect(detectMimeType('icon.png')).toBe('image/png');
    expect(detectMimeType('anim.gif')).toBe('image/gif');
    expect(detectMimeType('vector.svg')).toBe('image/svg+xml');
  });

  it('detects audio/video types', () => {
    expect(detectMimeType('song.mp3')).toBe('audio/mpeg');
    expect(detectMimeType('video.mp4')).toBe('video/mp4');
    expect(detectMimeType('clip.webm')).toBe('video/webm');
  });

  it('detects text types', () => {
    expect(detectMimeType('readme.md')).toBe('text/markdown');
    expect(detectMimeType('data.csv')).toBe('text/csv');
    expect(detectMimeType('config.json')).toBe('application/json');
  });

  it('is case-insensitive', () => {
    expect(detectMimeType('PHOTO.JPG')).toBe('image/jpeg');
    expect(detectMimeType('Doc.PDF')).toBe('application/pdf');
  });

  it('returns octet-stream for unknown extensions', () => {
    expect(detectMimeType('file.xyz')).toBe('application/octet-stream');
    expect(detectMimeType('noext')).toBe('application/octet-stream');
  });

  it('handles files with multiple dots', () => {
    expect(detectMimeType('archive.tar.gz')).toBe('application/gzip');
    expect(detectMimeType('my.file.name.pdf')).toBe('application/pdf');
  });
});
