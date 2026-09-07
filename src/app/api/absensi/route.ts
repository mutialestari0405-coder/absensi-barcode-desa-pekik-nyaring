import { NextRequest, NextResponse } from 'next/server'
import { db } from '@/lib/db'
import { getAuthUser } from '@/lib/auth'
import { getTanggalJakarta } from '@/lib/waktu'

// GET: riwayat absensi dengan filter tanggal & pencarian
export async function GET(req: NextRequest) {
  try {
    const user = getAuthUser(req)
    if (!user) {
      return NextResponse.json({ success: false, message: 'Tidak memiliki akses' }, { status: 401 })
    }

    const { searchParams } = new URL(req.url)
    const dari = searchParams.get('dari')
    const sampai = searchParams.get('sampai')
    const q = searchParams.get('q')?.trim()

    const today = getTanggalJakarta()
    const where: {
      tanggal?: { gte?: string; lte?: string }
      perangkat?: {
        OR?: Array<
          { nama?: { contains: string } } | { nipd?: { contains: string } } | { jabatan?: { contains: string } }
        >
      }
    } = {}

    if (dari || sampai) {
      where.tanggal = {}
      if (dari) where.tanggal.gte = dari
      if (sampai) where.tanggal.lte = sampai
    } else {
      where.tanggal = { gte: today, lte: today }
    }

    if (q) {
      where.perangkat = {
        OR: [
          { nama: { contains: q } },
          { nipd: { contains: q } },
          { jabatan: { contains: q } },
        ],
      }
    }

    const absensi = await db.absensi.findMany({
      where,
      include: { perangkat: true },
      orderBy: [{ tanggal: 'desc' }, { jamDatang: 'desc' }],
      take: 1000,
    })

    return NextResponse.json({ success: true, data: absensi })
  } catch {
    return NextResponse.json(
      { success: false, message: 'Terjadi kesalahan pada server' },
      { status: 500 }
    )
  }
}

const KETERANGAN_SAH = ['MASUK', 'IZIN', 'SAKIT', 'DINAS_LUAR', 'ALPA'] as const

// POST: catat absensi manual oleh admin — untuk perangkat yang tidak/harusnya
// tidak scan, misalnya Izin, Sakit, Dinas Luar, atau Alpa
export async function POST(req: NextRequest) {
  try {
    const user = getAuthUser(req)
    if (!user) {
      return NextResponse.json({ success: false, message: 'Tidak memiliki akses' }, { status: 401 })
    }

    const body = await req.json().catch(() => null)
    const perangkatId = Number(body?.perangkatId)
    const tanggal: string = typeof body?.tanggal === 'string' ? body.tanggal.trim() : ''
    const keterangan: string = typeof body?.keteranganKehadiran === 'string'
      ? body.keteranganKehadiran.trim().toUpperCase()
      : 'MASUK'

    if (!Number.isInteger(perangkatId) || !/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) {
      return NextResponse.json(
        { success: false, message: 'Perangkat dan tanggal (YYYY-MM-DD) wajib diisi' },
        { status: 400 }
      )
    }
    if (!KETERANGAN_SAH.includes(keterangan as never)) {
      return NextResponse.json(
        { success: false, message: 'Keterangan harus salah satu dari: Masuk, Izin, Sakit, Dinas Luar, Alpa' },
        { status: 400 }
      )
    }

    const perangkat = await db.perangkat.findUnique({ where: { id: perangkatId } })
    if (!perangkat) {
      return NextResponse.json({ success: false, message: 'Perangkat desa tidak ditemukan' }, { status: 404 })
    }

    // Jam opsional (HH:mm atau HH:mm:ss)
    const jamRegex = /^([01]\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/
    function rapikanJam(nilai: unknown): string | null {
      if (typeof nilai !== 'string' || nilai.trim() === '') return null
      if (!jamRegex.test(nilai.trim())) return null
      const [j, m, d = '00'] = nilai.trim().split(':')
      return `${j}:${m}:${d}`
    }
    const jamDatang = rapikanJam(body?.jamDatang)
    const jamPulang = rapikanJam(body?.jamPulang)

    const ada = await db.absensi.findUnique({
      where: { perangkatId_tanggal: { perangkatId, tanggal } },
    })
    if (ada) {
      return NextResponse.json(
        { success: false, message: `Sudah ada record absensi ${perangkat.nama} pada tanggal itu. Ubah keterangannya lewat tabel.` },
        { status: 409 }
      )
    }

    const dibuat = await db.absensi.create({
      data: {
        perangkatId,
        tanggal,
        jamDatang,
        jamPulang,
        kategori: jamDatang && jamPulang ? 'AKTIF' : 'PASIF',
        keteranganKehadiran: keterangan,
        metode: 'MANUAL',
      },
    })

    return NextResponse.json({
      success: true,
      message: `Absensi manual ${perangkat.nama} (${keterangan}) berhasil dicatat`,
      data: dibuat,
    })
  } catch {
    return NextResponse.json(
      { success: false, message: 'Gagal mencatat absensi manual' },
      { status: 500 }
    )
  }
}
