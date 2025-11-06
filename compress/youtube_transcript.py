#!/usr/bin/env python3
"""
YouTube Transcript Extractor
Extracts transcripts from YouTube videos in multiple formats
"""

import sys
import json
import argparse
from youtube_transcript_api import YouTubeTranscriptApi
from youtube_transcript_api._errors import TranscriptsDisabled, NoTranscriptFound, VideoUnavailable

def format_time(seconds):
    """Convert seconds to SRT/VTT time format (HH:MM:SS,mmm)"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"

def format_time_vtt(seconds):
    """Convert seconds to VTT time format (HH:MM:SS.mmm)"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}.{millis:03d}"

def get_transcript(video_id, language_codes=None):
    """Get transcript for a video, trying multiple languages"""
    try:
        # Try to get transcript in preferred languages first
        if language_codes:
            try:
                transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)
                for lang_code in language_codes:
                    try:
                        transcript = transcript_list.find_transcript([lang_code])
                        return transcript.fetch()
                    except:
                        continue
            except:
                pass
        
        # Fall back to auto-generated or available transcript
        transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)
        
        # Try to get manually created transcript first
        try:
            transcript = transcript_list.find_manually_created_transcript(['en'])
            return transcript.fetch()
        except:
            pass
        
        # Try auto-generated English
        try:
            transcript = transcript_list.find_generated_transcript(['en'])
            return transcript.fetch()
        except:
            pass
        
        # Get any available transcript
        for transcript_item in transcript_list:
            try:
                return transcript_item.fetch()
            except:
                continue
        
        # If we get here, no transcript was found
        raise Exception("No transcript available for this video")
        
    except TranscriptsDisabled:
        raise Exception("Transcripts are disabled for this video")
    except NoTranscriptFound:
        raise Exception("No transcript found for this video")
    except VideoUnavailable:
        raise Exception("Video is unavailable or doesn't exist")
    except Exception as e:
        raise Exception(f"Error fetching transcript: {str(e)}")

def format_as_text(transcript_data, include_timestamps=False):
    """Format transcript as plain text"""
    if include_timestamps:
        lines = []
        for item in transcript_data:
            start_time = format_time(item['start']).replace(',', ':')
            text = item['text'].strip()
            lines.append(f"[{start_time}] {text}")
        return '\n'.join(lines)
    else:
        return '\n'.join([item['text'].strip() for item in transcript_data])

def format_as_json(transcript_data):
    """Format transcript as JSON"""
    return json.dumps(transcript_data, indent=2, ensure_ascii=False)

def format_as_srt(transcript_data):
    """Format transcript as SRT subtitle format"""
    srt_lines = []
    for index, item in enumerate(transcript_data, start=1):
        start_time = format_time(item['start'])
        end_time = format_time(item['start'] + item.get('duration', 3.0))
        text = item['text'].strip().replace('\n', ' ')
        srt_lines.append(f"{index}\n{start_time} --> {end_time}\n{text}\n")
    return '\n'.join(srt_lines)

def format_as_vtt(transcript_data):
    """Format transcript as VTT subtitle format"""
    vtt_lines = ["WEBVTT", ""]
    for item in transcript_data:
        start_time = format_time_vtt(item['start'])
        end_time = format_time_vtt(item['start'] + item.get('duration', 3.0))
        text = item['text'].strip().replace('\n', ' ')
        vtt_lines.append(f"{start_time} --> {end_time}\n{text}")
    return '\n'.join(vtt_lines)

def main():
    parser = argparse.ArgumentParser(description='Extract YouTube video transcript')
    parser.add_argument('video_id', help='YouTube video ID')
    parser.add_argument('--format', choices=['txt', 'txt-timestamps', 'json', 'srt', 'vtt'], 
                       default='txt', help='Output format')
    parser.add_argument('--language', help='Language code (e.g., en, es, fr)', default='en')
    
    args = parser.parse_args()
    
    try:
        # Get transcript
        transcript_data = get_transcript(args.video_id, [args.language, 'en'])
        
        # Format based on requested format
        if args.format == 'txt':
            output = format_as_text(transcript_data, include_timestamps=False)
            content_type = 'text/plain'
        elif args.format == 'txt-timestamps':
            output = format_as_text(transcript_data, include_timestamps=True)
            content_type = 'text/plain'
        elif args.format == 'json':
            output = format_as_json(transcript_data)
            content_type = 'application/json'
        elif args.format == 'srt':
            output = format_as_srt(transcript_data)
            content_type = 'text/srt'
        elif args.format == 'vtt':
            output = format_as_vtt(transcript_data)
            content_type = 'text/vtt'
        
        # Return as JSON for API
        result = {
            'success': True,
            'format': args.format,
            'content': output,
            'content_type': content_type,
            'video_id': args.video_id,
            'entries_count': len(transcript_data)
        }
        
        print(json.dumps(result))
        
    except Exception as e:
        error_result = {
            'success': False,
            'error': str(e),
            'video_id': args.video_id
        }
        print(json.dumps(error_result))
        sys.exit(1)

if __name__ == '__main__':
    main()

