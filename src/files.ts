import fs from 'node:fs';
import path from 'node:path';
import type { Writable } from 'node:stream';

export interface SaveOptions {
  /** Destination file (`--out`); omitted -> write to stdout. */
  out?: string;
  /** Overwrite an existing `out` (`--force`). */
  force?: boolean;
}

export interface SaveResult {
  /** Absolute-or-as-given path written; absent when streamed to stdout or nothing was written (dry-run). */
  path?: string;
  bytes: number;
}

type Source = ReadableStream<Uint8Array> | Uint8Array | null;

/**
 * Write raw bytes (a `requestBinary` body or a Buffer) to `opts.out`, or to `stdout` when no path is
 * given. Never decodes. The file is written to a temp sibling and moved into place only after the
 * last byte arrived, so a failed download leaves no partial file and never clobbers an existing one.
 * Without `force` an existing `out` is an error (`File exists: … Use --force to overwrite.`), checked
 * before reading the source and again atomically at the final link. `null` (dry-run) writes nothing.
 */
export async function saveOrStream(source: Source, opts: SaveOptions = {}, stdout: Writable = process.stdout): Promise<SaveResult> {
  const { out, force = false } = opts;
  if (out && !force && fs.existsSync(out)) throw existsError(out);
  if (source === null) return { bytes: 0 };

  if (!out) {
    let bytes = 0;
    for await (const chunk of chunks(source)) {
      await writeTo(stdout, chunk);
      bytes += chunk.byteLength;
    }
    return { bytes };
  }

  const tmp = path.join(path.dirname(out), `.${path.basename(out)}.${process.pid}.${Date.now()}.part`);
  let bytes = 0;
  try {
    const fh = await fs.promises.open(tmp, 'wx');
    try {
      for await (const chunk of chunks(source)) {
        await fh.write(chunk);
        bytes += chunk.byteLength;
      }
    } finally {
      await fh.close();
    }
    if (force) {
      fs.renameSync(tmp, out);
    } else {
      // link fails with EEXIST if `out` appeared meanwhile: no-overwrite stays atomic.
      try {
        fs.linkSync(tmp, out);
      } catch (err: unknown) {
        if (err instanceof Error && 'code' in err && err.code === 'EEXIST') throw existsError(out);
        throw err;
      }
      fs.rmSync(tmp, { force: true });
    }
  } catch (err: unknown) {
    fs.rmSync(tmp, { force: true });
    throw err;
  }
  return { path: out, bytes };
}

function existsError(out: string): Error {
  return new Error(`File exists: ${out}. Use --force to overwrite.`);
}

async function* chunks(source: ReadableStream<Uint8Array> | Uint8Array): AsyncGenerator<Uint8Array> {
  if (source instanceof Uint8Array) {
    if (source.byteLength > 0) yield source;
    return;
  }
  const reader = source.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) return;
      if (value.byteLength > 0) yield value;
    }
  } finally {
    reader.releaseLock();
  }
}

/** Write one chunk honouring backpressure; rejects on a stream error. Does not end the stream. */
function writeTo(stream: Writable, chunk: Uint8Array): Promise<void> {
  const { promise, resolve, reject } = Promise.withResolvers<void>();
  stream.write(chunk, (err) => (err ? reject(err) : resolve()));
  return promise;
}
