import DaVinciResolveScript as dvr
import argparse
from pathlib import Path
import os
import re
from datetime import datetime


def _read_text(path: Path):
    encodings = ['utf-8', 'utf-8-sig', 'gb18030', 'latin-1']
    for enc in encodings:
        try:
            with path.open('r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
    with path.open('r', errors='ignore') as f:
        return f.read()


def parse_srt_for_timestamp(srt_path: Path):
    content = _read_text(srt_path)
    m = re.search(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3})", content)
    if not m:
        return None
    return datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S.%f")


def get_clip_duration_frames(srt_path: Path, fps: int):
    content = _read_text(srt_path)
    diff_times = re.findall(r"DiffTime: (\d+)ms", content)
    if diff_times:
        total_ms = sum(int(x) for x in diff_times)
        frame_ms = 1000.0 / fps
        return int(total_ms / frame_ms)
    entries = re.findall(r"FrameCnt: (\d+)", content)
    if entries:
        return len(entries)
    return None


def compute_record_frame(ts: datetime, fps: int):
    day_start = ts.replace(hour=0, minute=0, second=0, microsecond=0)
    delta = ts - day_start
    return int(round(delta.total_seconds() * fps))


def list_cam_folders(date_folder: Path):
    cams = []
    for p in sorted(date_folder.iterdir()):
        if p.is_dir() and p.name.upper().startswith("CAM"):
            cams.append(p)
    return cams


def _select_indices(count: int, prompt: str):
    sel = input(prompt).strip()
    result = []
    if sel:
        for token in re.split(r"[，,\s]+", sel):
            if not token:
                continue
            try:
                i = int(token)
                if 1 <= i <= count:
                    if i not in result:
                        result.append(i)
            except Exception:
                continue
    return result


def norm_ext(name: str):
    lower = name.lower()
    if lower.endswith('.mp4') or lower.endswith('.mov'):
        return 'video'
    if lower.endswith('.srt'):
        return 'srt'
    return ''


def find_srt_for_video(video_path: Path):
    base = video_path.with_suffix('')
    candidate_lower = base.with_suffix('.srt')
    candidate_upper = base.with_suffix('.SRT')
    if candidate_lower.exists():
        if not candidate_lower.name.startswith('._') and not candidate_lower.name.startswith('.'):
            return candidate_lower
    if candidate_upper.exists():
        if not candidate_upper.name.startswith('._') and not candidate_upper.name.startswith('.'):
            return candidate_upper
    for srt in video_path.parent.glob('*.srt'):
        if srt.stem == video_path.stem:
            if not srt.name.startswith('._') and not srt.name.startswith('.'):
                return srt
    for srt in video_path.parent.glob('*.SRT'):
        if srt.stem == video_path.stem:
            if not srt.name.startswith('._') and not srt.name.startswith('.'):
                return srt
    return None


def scan_videos(cam_folder: Path):
    videos = []
    for p in cam_folder.rglob('*'):
        if p.name.startswith('._') or p.name.startswith('.'):
            continue
        if p.is_file() and norm_ext(p.name) == 'video':
            videos.append(p)
    return videos


def ensure_bin(media_pool, root_folder, name: str):
    if root_folder is None:
        return None
    subs = None
    try:
        subs = root_folder.GetSubFolders()
    except Exception:
        subs = None
    if isinstance(subs, dict):
        for f in subs.values():
            try:
                if f.GetName() == name:
                    return f
            except Exception:
                continue
    try:
        media_pool.SetCurrentFolder(root_folder)
    except Exception:
        pass
    created = media_pool.AddSubFolder(root_folder, name)
    if created:
        return created
    # Fallback: re-scan
    try:
        subs = root_folder.GetSubFolders()
        if isinstance(subs, dict):
            for f in subs.values():
                try:
                    if f.GetName() == name:
                        return f
                except Exception:
                    continue
    except Exception:
        pass
    return None


def get_or_create_timeline(project, media_pool, name: str):
    cnt = project.GetTimelineCount()
    for i in range(1, cnt + 1):
        tl = project.GetTimelineByIndex(i)
        if tl and tl.GetName() == name:
            return tl
    return media_pool.CreateEmptyTimeline(name)


def safe_filename(name: str) -> str:
    return re.sub(r"[\\/:*?\"<>|]", "_", name)


def main():
    parser = argparse.ArgumentParser(description='同步无人机素材并按日期/机位生成时间线与XML')
    parser.add_argument('-p', '--path', required=True, help='待处理的日期文件夹路径或其上级路径')
    parser.add_argument('-f', '--frame_rate', required=True, type=int, help='达芬奇项目帧率')
    parser.add_argument('-o', '--output', required=True, help='导出XML的目标文件夹路径')
    args = parser.parse_args()

    date_folder = Path(args.path).resolve()
    export_folder = Path(args.output).resolve()
    fps = int(args.frame_rate)

    if not date_folder.exists() or not date_folder.is_dir():
        print('日期文件夹不存在')
        return
    export_folder.mkdir(parents=True, exist_ok=True)
    if fps <= 0:
        print('帧率必须为正整数')
        return

    cams = list_cam_folders(date_folder)
    if not cams:
        subdirs = [p for p in sorted(date_folder.iterdir()) if p.is_dir()]
        if not subdirs:
            print('未检测到机位文件夹或日期子文件夹')
            return
        for idx, d in enumerate(subdirs, start=1):
            print(f"{idx}. {d.name}")
        idxs = _select_indices(len(subdirs), '请选择一个日期文件夹索引：')
        if not idxs:
            print('未选择日期文件夹')
            return
        i = idxs[0]
        date_folder = subdirs[i - 1]
        cams = list_cam_folders(date_folder)
        if not cams:
            print('所选日期文件夹下仍未检测到机位文件夹（CAM*）')
            return

    for idx, cam in enumerate(cams, start=1):
        print(f"{idx}. {cam.name}")
    cam_idxs = _select_indices(len(cams), '请输入待处理机位的索引（支持逗号分隔多个）：')
    if not cam_idxs:
        print('未选择任何机位')
        return
    chosen_cams = [cams[i - 1] for i in cam_idxs]

    resolve = dvr.scriptapp('Resolve')
    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject()
    if not project:
        print('未检测到打开的达芬奇项目，请先打开 DaVinci Resolve 并加载项目')
        return
    media_pool = project.GetMediaPool()
    root = media_pool.GetRootFolder()
    project.SetSetting('timelineFrameRate', str(fps))

    all_timelines = []
    for cam_folder in chosen_cams:
        cam_id = cam_folder.name
        cam_bin = ensure_bin(media_pool, root, cam_id)
        if cam_bin is None:
            print(f'创建或获取机位文件夹失败: {cam_id}')
            continue
        videos = scan_videos(cam_folder)
        group_by_date = {}
        for v in videos:
            srt = find_srt_for_video(v)
            if not srt:
                print(f'跳过无SRT的素材: {v}')
                continue
            ts = parse_srt_for_timestamp(srt)
            if not ts:
                print(f'跳过无法解析时间戳的素材: {v}')
                continue
            realdate = ts.strftime('%y-%m-%d')
            group_by_date.setdefault(realdate, []).append((v, srt, ts))

        for realdate, items in sorted(group_by_date.items()):
            date_bin = ensure_bin(media_pool, cam_bin, realdate)
            media_pool.SetCurrentFolder(date_bin)
            if date_bin is None:
                print(f'创建或获取日期文件夹失败: {cam_id}/{realdate}')
                continue
            imported_items = []
            for v, srt, ts in items:
                added = media_pool.ImportMedia([str(v)])
                if added:
                    imported_items.append((added[0], v, srt, ts))
            timeline_name = f"{date_folder.name}_{cam_id}_{realdate}"
            timeline = get_or_create_timeline(project, media_pool, timeline_name)
            if not timeline:
                print(f'创建时间线失败: {timeline_name}')
                continue
            timeline.SetSetting('timelineFrameRate', str(fps))
            timeline.SetStartTimecode('00:00:00:00')
            max_w, max_h = 0, 0
            clip_infos = []
            for item, v, srt, ts in imported_items:
                try:
                    res = item.GetClipProperty('Resolution')
                    if isinstance(res, str):
                        m = re.search(r"(\d{3,5})x(\d{3,5})", res)
                        if m:
                            w, h = int(m.group(1)), int(m.group(2))
                            if w * h > max_w * max_h:
                                max_w, max_h = w, h
                    elif isinstance(res, dict):
                        val = res.get('Resolution')
                        if isinstance(val, str):
                            m = re.search(r"(\d{3,5})x(\d{3,5})", val)
                            if m:
                                w, h = int(m.group(1)), int(m.group(2))
                                if w * h > max_w * max_h:
                                    max_w, max_h = w, h
                except Exception:
                    pass
                record_frame = compute_record_frame(ts, fps)
                duration = get_clip_duration_frames(srt, fps)
                info = {
                    'mediaPoolItem': item,
                    'trackIndex': 1,
                    'recordFrame': record_frame,
                }
                if duration is not None:
                    info['startFrame'] = 0
                    info['endFrame'] = duration
                clip_infos.append(info)
            if max_w > 0 and max_h > 0:
                timeline.SetSetting('useCustomSettings', '1')
                timeline.SetSetting('timelineResolutionWidth', str(max_w))
                timeline.SetSetting('timelineResolutionHeight', str(max_h))
                rw = timeline.GetSetting('timelineResolutionWidth')
                rh = timeline.GetSetting('timelineResolutionHeight')
                print(f"Timeline '{timeline_name}' 分辨率: {rw}x{rh}")
            clip_infos.sort(key=lambda x: x['recordFrame'])
            for info in clip_infos:
                ok = media_pool.AppendToTimeline([info])
                if not ok:
                    print(f"追加失败: {info['mediaPoolItem'].GetName()} @ {info['recordFrame']}")
            all_timelines.append(timeline)

    for tl in all_timelines:
        xml_name = f"{safe_filename(tl.GetName())}.xml"
        xml_path = str(export_folder / xml_name)
        try:
            p = Path(xml_path)
            if p.exists():
                p.unlink()
        except Exception:
            pass
        project.SetCurrentTimeline(tl)
        ok = tl.Export(xml_path, resolve.EXPORT_FCP_7_XML, resolve.EXPORT_NONE)
        if ok:
            print(f'已导出: {xml_path}')
        else:
            print(f'导出失败: {xml_path}')


if __name__ == '__main__':
    main()