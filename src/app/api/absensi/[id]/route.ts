import { NextRequest, NextResponse } from 'next/server'
import { db } from '@/lib/db'
import { getAuthUser } from '@/lib/auth'

// Nilai keterangan kehadiran yang sah
const KETERANGAN_SAH = ['MASUK', 'IZIN', 'SAKIT', 'DINAS_LUAR', 'ALPA'] as const

// PATCH: ubah keterangan kehadiran (Masuk / Izin / Sakit / Dinas Luar / Alpa)
// serta jam datang & pulang (opsional) pada satu record absensi
export async function PATCH(req: NextRequest, ctx: { params: Promise<{ id: string }> }) {
  try {
    const user = getAuthUser(req)
    if (!user) {
      return NextResponse.json({ success: false, message: 'Tidak memiliki akses' }, { status: 401 })
    }

    const { id } = await ctx.params
    const absensiId = Number(id)
    if (!Number.isInteger(absensiId)) {
      return NextResponse.json({ success: false, message: 'ID tidak valid' }, { status: 400 })
    }

    const body = await req.json().catch(() => null)
    const keterangan: unknown = body?.keteranganKehadiran
    if (typeof keterangan !== 'string' || !KETERANGAN_SAH.includes(keterangan as never)) {
      return NextResponse.json(
        { success: false, message: 'Keterangan harus salah satu dari: Masuk, Izin, Sakit, Dinas Luar, Alpa' },
        { status: 400 }
      )
    }

    const existing = await db.absensi.findUnique({ where: { id: absensiId } })
    if (!existing) {
      return NextResponse.json({ success: false, message: 'Data absensi tidak ditemukan' }, { status: 404 })
    }

    const data: { keteranganKehadiran: string } = { keteranganKehadiran: keterangan }

    // Jam opsional — boleh dikoreksi admin bersamaan dengan keterangan
    const jamRegex = /^([01]\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/
    if ('jamDatang' in (body ?? {})) {
      const jd = body.jamDatang
      if (typeof jd === 'string' && jd.trim() !== '') {
        if (!jamRegex.test(jd.trim())) {
          return NextResponse.json({ success: false, message: 'Format jam datang salah (HH:mm)' }, { status: 400 })
        }
        const [j, m, d = '00'] = jd.trim().split(':')
        Object.assign(data, { jamDatang: `${j}:${m}:${d}` })
      } else if (jd === null) {
        Object.assign(data, { jamDatang: null })
      }
    }
    if ('jamPulang' in (body ?? {})) {
      const jp = body.jamPulang
      if (typeof jp === 'string' && jp.trim() !== '') {
        if (!jamRegex.test(jp.trim())) {
          return NextResponse.json({ success: false, message: 'Format jam pulang salah (HH:mm)' }, { status: 400 })
        }
        const [j, m, d = '00'] = jp.trim().split(':')
        Object.assign(data, { jamPulang: `${j}:${m}:${d}` })
      } else if (jp === null) {
        Object.assign(data, { jamPulang: null })
      }
    }

    const updated = await db.absensi.update({ where: { id: absensiId }, data })

    // Perbarui kategori otomatis: AKTIF bila ada jam datang & pulang
    const kategori = updated.jamDatang && updated.jamPulang ? 'AKTIF' : 'PASIF'
    const final = await db.absensi.update({ where: { id: absensiId }, data: { kategori } })

    return NextResponse.json({
      success: true,
      message: 'Keterangan absensi berhasil diperbarui',
      data: final,
    })
  } catch {
    return NextResponse.json(
      { success: false, message: 'Gagal memperbarui data absensi' },
      { status: 500 }
    )
  }
}

// DELETE: hapus record absensi (koreksi oleh admin)
export async function DELETE(req: NextRequest, ctx: { params: Promise<{ id: string }> }) {
  try {
    const user = getAuthUser(req)
    if (!user) {
      return NextResponse.json({ success: false, message: 'Tidak memiliki akses' }, { status: 401 })
    }

    const { id } = await ctx.params
    const absensiId = Number(id)

    const existing = await db.absensi.findUnique({ where: { id: absensiId } })
    if (!existing) {
      return NextResponse.json({ success: false, message: 'Data absensi tidak ditemukan' }, { status: 404 })
    }

    await db.absensi.delete({ where: { id: absensiId } })
    return NextResponse.json({ success: true, message: 'Data absensi berhasil dihapus' })
  } catch {
    return NextResponse.json({ success: false, message: 'Gagal menghapus data absensi' }, { status: 500 })
  }
}
