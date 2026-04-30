import { ref, onMounted, onUnmounted } from 'vue'

export interface UseDropOverlayOptions {
  isActive: () => boolean
  onFiles: (files: File[]) => void
}

export function useDropOverlay(options: UseDropOverlayOptions) {
  const showOverlay = ref(false)
  let dragCounter = 0

  function onDocDragenter(e: DragEvent) {
    if (!options.isActive()) return
    e.preventDefault()
    if (++dragCounter === 1) showOverlay.value = true
  }

  function onDocDragleave() {
    if (dragCounter === 0) return
    if (--dragCounter <= 0) {
      dragCounter = 0
      showOverlay.value = false
    }
  }

  function onDocDragover(e: DragEvent) {
    if (showOverlay.value) e.preventDefault()
  }

  function onDocDrop(e: DragEvent) {
    e.preventDefault()
    dragCounter = 0
    showOverlay.value = false
  }

  function onOverlayDrop(e: DragEvent) {
    e.preventDefault()
    dragCounter = 0
    showOverlay.value = false
    const dt = e.dataTransfer
    if (!dt) return
    const files: File[] = []
    if (dt.items) {
      for (const item of Array.from(dt.items)) {
        if (item.kind !== 'file') continue
        const entry = item.webkitGetAsEntry?.()
        if (entry?.isDirectory) continue
        const f = item.getAsFile()
        if (f) files.push(f)
      }
    } else {
      files.push(...Array.from(dt.files))
    }
    if (files.length) options.onFiles(files)
  }

  onMounted(() => {
    document.addEventListener('dragenter', onDocDragenter)
    document.addEventListener('dragleave', onDocDragleave)
    document.addEventListener('dragover', onDocDragover)
    document.addEventListener('drop', onDocDrop)
  })

  onUnmounted(() => {
    document.removeEventListener('dragenter', onDocDragenter)
    document.removeEventListener('dragleave', onDocDragleave)
    document.removeEventListener('dragover', onDocDragover)
    document.removeEventListener('drop', onDocDrop)
  })

  return {
    showOverlay,
    onOverlayDrop,
  }
}
