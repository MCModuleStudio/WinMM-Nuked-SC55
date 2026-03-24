#include <windows.h>
#include <mmsystem.h>
#include "audio.h"
#include "cast.h"
#include "config.h"
#include "emu.h"
#include "math_util.h"
#include <algorithm>
#include <cinttypes>
#include <condition_variable>
#include <cstddef>
#include <cstdio>
#include <cstring>
#include <functional>
#include <memory>
#include <mutex>
#include <source_location>
#include <string>
#include <thread>

typedef UINT     (WINAPI *midiOutGetNumDevsProc)     (void);
typedef MMRESULT (WINAPI *midiOutGetDevCapsWProc)    (UINT_PTR uDeviceID, LPMIDIOUTCAPSW pmoc, UINT cboc);
typedef MMRESULT (WINAPI *midiOutOpenProc)           (LPHMIDIOUT phmo, UINT uDeviceID, DWORD_PTR dwCallback, DWORD_PTR dwInstance, DWORD fdwOpen);
typedef MMRESULT (WINAPI *midiOutCloseProc)          (HMIDIOUT hmo);
typedef MMRESULT (WINAPI *midiOutShortMsgProc)       (HMIDIOUT hmo, DWORD dwMsg);
typedef MMRESULT (WINAPI *midiOutLongMsgProc)        (HMIDIOUT hmo, LPMIDIHDR pmh, UINT cbmh);
typedef MMRESULT (WINAPI *midiOutMessageProc)        (HMIDIOUT hmo, UINT uMsg, DWORD_PTR dw1, DWORD_PTR dw2);
typedef MMRESULT (WINAPI *midiOutResetProc)          (HMIDIOUT hmo);
typedef MMRESULT (WINAPI *midiOutSetVolumeProc)      (HMIDIOUT hmo, DWORD dwVolume);
typedef MMRESULT (WINAPI *waveOutOpenProc)           (LPHWAVEOUT, UINT, LPCWAVEFORMATEX, DWORD_PTR, DWORD_PTR, DWORD);
typedef MMRESULT (WINAPI *waveOutCloseProc)          (HWAVEOUT);
typedef MMRESULT (WINAPI *waveOutPrepareHeaderProc)  (HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT (WINAPI *waveOutUnprepareHeaderProc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT (WINAPI *waveOutWriteProc)          (HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT (WINAPI *waveOutResetProc)          (HWAVEOUT);
typedef MMRESULT (WINAPI *waveOutGetErrorTextAProc)  (MMRESULT, LPSTR, UINT);

HMODULE hWinMM = nullptr;
midiOutGetNumDevsProc      pmidiOutGetNumDevs      = nullptr;
midiOutGetDevCapsWProc     pmidiOutGetDevCapsW     = nullptr;
midiOutOpenProc            pmidiOutOpen            = nullptr;
midiOutCloseProc           pmidiOutClose           = nullptr;
midiOutShortMsgProc        pmidiOutShortMsg        = nullptr;
midiOutLongMsgProc         pmidiOutLongMsg         = nullptr;
midiOutMessageProc         pmidiOutMessage         = nullptr;
midiOutResetProc           pmidiOutReset           = nullptr;
midiOutSetVolumeProc       pmidiOutSetVolume       = nullptr;
waveOutOpenProc            pwaveOutOpen            = nullptr;
waveOutCloseProc           pwaveOutClose           = nullptr;
waveOutPrepareHeaderProc   pwaveOutPrepareHeader   = nullptr;
waveOutUnprepareHeaderProc pwaveOutUnprepareHeader = nullptr;
waveOutWriteProc           pwaveOutWrite           = nullptr;
waveOutResetProc           pwaveOutReset           = nullptr;
waveOutGetErrorTextAProc   pwaveOutGetErrorTextA   = nullptr;

// I don't know C++ BTW
// So codes in this namespace are vibe coded
namespace {

struct WinmmConfig {
    std::filesystem::path rom_dir;
    std::filesystem::path nvram_file;
    Romset romset = Romset::MK2;
    bool detect_by_filename = true;
    bool detect_by_hash = true;
    EMU_SystemReset system_reset = EMU_SystemReset::NONE;
    uint32_t audio_buffer_frames = 1024;
};

struct WinmmAudioOutput {
    static constexpr size_t kBufferCount = 4;

    HWAVEOUT device = nullptr;
    HANDLE done_event = nullptr;
    uint32_t sample_rate = 0;
    uint32_t buffer_frames = 0;

    std::array<std::vector<int16_t>, kBufferCount> buffers{};
    std::array<WAVEHDR, kBufferCount> headers{};
    std::array<bool, kBufferCount> submitted{};

    size_t active_buffer = 0;
    size_t write_pos = 0; // samples (interleaved LR)
    uint16_t lvolume = 0xFFFF, rvolume = 0xFFFF;

    static void LogWaveError(const char* op, MMRESULT result) {
        char text[MAXERRORLENGTH] = {};
        pwaveOutGetErrorTextA(result, text, MAXERRORLENGTH);
        std::fprintf(stderr, "[WinMM] %s failed: %s (%u)\n", op, text, (unsigned)result);
    }

    bool IsBufferFree(size_t index) {
        if (!submitted[index]) {
            return true;
        }

        if (headers[index].dwFlags & WHDR_DONE) {
            submitted[index] = false;
            return true;
        }

        return false;
    }

    void WaitUntilBufferFree(size_t index) {
        while (!IsBufferFree(index)) {
            if (done_event) {
                (void)WaitForSingleObject(done_event, INFINITE);
            } else {
                Sleep(1);
            }
        }
    }

    bool SubmitBuffer(size_t index) {
        WAVEHDR& hdr = headers[index];
        hdr.dwFlags &= ~WHDR_DONE;
        hdr.dwBytesRecorded = static_cast<DWORD>(buffers[index].size() * sizeof(int16_t));

        MMRESULT write = pwaveOutWrite(device, &hdr, sizeof(hdr));
        if (write != MMSYSERR_NOERROR) {
            LogWaveError("waveOutWrite", write);
            return false;
        }

        submitted[index] = true;
        return true;
    }

    void OnSample(const AudioFrame<int32_t>& frame) {
        if (!device) {
            return;
        }

        AudioFrame<int16_t> out{};
        Normalize(frame, out);

        auto& buf = buffers[active_buffer];
        if (write_pos + 1 >= buf.size()) {
            return;
        }

        buf[write_pos++] = (int16_t) (((int32_t) out.left  * (int32_t) lvolume) >> 16);
        buf[write_pos++] = (int16_t) (((int32_t) out.right * (int32_t) rvolume) >> 16);

        if (write_pos < buf.size()) {
            return;
        }

        if (!SubmitBuffer(active_buffer)) {
            return;
        }

        const size_t next = (active_buffer + 1) % kBufferCount;
        WaitUntilBufferFree(next);

        active_buffer = next;
        write_pos = 0;
    }

    bool Init(uint32_t rate, uint32_t frames) {
        Shutdown();

        sample_rate = rate;
        buffer_frames = frames;

        done_event = CreateEventW(nullptr, FALSE, FALSE, nullptr);
        if (!done_event) {
            std::fprintf(stderr, "[WinMM] CreateEventW failed\n");
            return false;
        }

        WAVEFORMATEX fmt{};
        fmt.wFormatTag = WAVE_FORMAT_PCM;
        fmt.nChannels = 2;
        fmt.nSamplesPerSec = sample_rate;
        fmt.wBitsPerSample = 16;
        fmt.nBlockAlign = static_cast<WORD>(fmt.nChannels * (fmt.wBitsPerSample / 8));
        fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * fmt.nBlockAlign;

        const MMRESULT open = pwaveOutOpen(
            &device,
            WAVE_MAPPER,
            &fmt,
            reinterpret_cast<DWORD_PTR>(done_event),
            0,
            CALLBACK_EVENT
        );
        if (open != MMSYSERR_NOERROR) {
            LogWaveError("waveOutOpen", open);
            device = nullptr;
            CloseHandle(done_event);
            done_event = nullptr;
            return false;
        }

        const size_t samples_per_buffer = static_cast<size_t>(buffer_frames) * AudioFrame<int16_t>::channel_count;
        for (size_t i = 0; i < kBufferCount; ++i) {
            buffers[i].assign(samples_per_buffer, 0);
            headers[i] = {};
            headers[i].lpData = reinterpret_cast<LPSTR>(buffers[i].data());
            headers[i].dwBufferLength = static_cast<DWORD>(samples_per_buffer * sizeof(int16_t));
            submitted[i] = false;

            const MMRESULT prep = pwaveOutPrepareHeader(device, &headers[i], sizeof(headers[i]));
            if (prep != MMSYSERR_NOERROR) {
                LogWaveError("waveOutPrepareHeader", prep);
                Shutdown();
                return false;
            }
        }

        active_buffer = 0;
        write_pos = 0;

        return true;
    }

    void Shutdown() {
        if (!device && !done_event) {
            return;
        }

        if (device) {
            pwaveOutReset(device);

            for (size_t i = 0; i < kBufferCount; ++i) {
                if (headers[i].dwFlags & WHDR_PREPARED) {
                    (void)pwaveOutUnprepareHeader(device, &headers[i], sizeof(headers[i]));
                }
                headers[i] = {};
                submitted[i] = false;
                buffers[i].clear();
            }

            (void)pwaveOutClose(device);
            device = nullptr;
        }

        if (done_event) {
            CloseHandle(done_event);
            done_event = nullptr;
        }

        active_buffer = 0;
        write_pos = 0;
        buffer_frames = 0;
        sample_rate = 0;
    }
};

HINSTANCE g_module = nullptr;
std::unique_ptr<Emulator> g_emulator;
WinmmAudioOutput g_audio;
std::thread g_worker_thread;
std::atomic_bool g_worker_running = false;
bool g_emulator_ready = false;
std::mutex g_runtime_mutex;

void WinmmSampleCallback(void* userdata, const AudioFrame<int32_t>& frame) {
    auto* out = static_cast<WinmmAudioOutput*>(userdata);
    out->OnSample(frame);
}

std::filesystem::path GetModulePath(HINSTANCE module) {
    if (!module) {
        return {};
    }

    wchar_t buffer[MAX_PATH] = {};
    const DWORD len = GetModuleFileNameW(module, buffer, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        return {};
    }
    return std::filesystem::path(buffer);
}

std::wstring ReadIniString(const std::filesystem::path& ini,
                           const wchar_t* section,
                           const wchar_t* key,
                           const wchar_t* default_value) {
    wchar_t buffer[1024] = {};
    GetPrivateProfileStringW(section, key, default_value, buffer, 1024, ini.c_str());
    return std::wstring(buffer);
}

int ReadIniInt(const std::filesystem::path& ini,
               const wchar_t* section,
               const wchar_t* key,
               int default_value) {
    return GetPrivateProfileIntW(section, key, default_value, ini.c_str());
}

void SetupLogOutputFromConfig(HINSTANCE module) {
    const std::filesystem::path module_path = GetModulePath(module);
    const std::filesystem::path module_dir = module_path.empty() ? std::filesystem::current_path() : module_path.parent_path();
    const std::filesystem::path ini_path = module_dir / "nuked-sc55-winmm.ini";

    const std::wstring log_path = ReadIniString(ini_path, L"nuked-sc55", L"log_path", L"console");
    const bool show_console = ReadIniInt(ini_path, L"nuked-sc55", L"show_console", 0) != 0;

    if (_wcsicmp(log_path.c_str(), L"console") == 0) {
        if (!show_console) {
            return;
        }

        if (!GetConsoleWindow() && !AllocConsole()) {
            return;
        }

        FILE* stream = nullptr;
        (void)freopen_s(&stream, "CONOUT$", "w", stdout);
        (void)freopen_s(&stream, "CONOUT$", "w", stderr);
        (void)freopen_s(&stream, "CONIN$", "r", stdin);
        return;
    }

    std::filesystem::path target_path = std::filesystem::path(log_path);
    if (target_path.empty()) {
        target_path = module_dir / "nuked-sc55-winmm.log";
    } else if (!target_path.is_absolute()) {
        target_path = module_dir / target_path;
    }

    const std::filesystem::path parent = target_path.parent_path();
    if (!parent.empty()) {
        std::error_code ec;
        (void)std::filesystem::create_directories(parent, ec);
    }

    FILE* stream = nullptr;
    (void)_wfreopen_s(&stream, target_path.c_str(), L"a", stdout);
    (void)_wfreopen_s(&stream, target_path.c_str(), L"a", stderr);
}

std::string ToAsciiLower(std::wstring value) {
    std::string out;
    out.reserve(value.size());
    for (wchar_t ch : value) {
        if (ch > 0x7F) {
            return {};
        }
        out.push_back(static_cast<char>(std::tolower(static_cast<unsigned char>(ch))));
    }
    return out;
}

std::filesystem::path MakePathAbsolute(const std::filesystem::path& base,
                                       const std::filesystem::path& candidate) {
    if (candidate.empty()) {
        return {};
    }
    if (candidate.is_absolute()) {
        return candidate;
    }
    return base / candidate;
}

WinmmConfig LoadConfig(HINSTANCE module) {
    WinmmConfig cfg;

    const std::filesystem::path module_path = GetModulePath(module);
    const std::filesystem::path module_dir = module_path.empty() ? std::filesystem::current_path() : module_path.parent_path();
    const std::filesystem::path ini_path = module_dir / "nuked-sc55-winmm.ini";

    cfg.rom_dir = module_dir / "roms";

    const std::wstring rom_dir = ReadIniString(ini_path, L"nuked-sc55", L"rom_dir", L"");
    if (!rom_dir.empty()) {
        cfg.rom_dir = MakePathAbsolute(module_dir, std::filesystem::path(rom_dir));
    }

    const std::wstring nvram_file = ReadIniString(ini_path, L"nuked-sc55", L"nvram_file", L"");
    if (!nvram_file.empty()) {
        cfg.nvram_file = MakePathAbsolute(module_dir, std::filesystem::path(nvram_file));
    }

    Romset parsed = Romset::MK2;
    const std::string romset_str = ToAsciiLower(ReadIniString(ini_path, L"nuked-sc55", L"romset", L"mk2"));
    if (ParseRomsetName(romset_str, parsed)) {
        cfg.romset = parsed;
    }

    cfg.detect_by_filename = ReadIniInt(ini_path, L"nuked-sc55", L"detect_by_filename", 1) != 0;
    cfg.detect_by_hash = ReadIniInt(ini_path, L"nuked-sc55", L"detect_by_hash", 1) != 0;

    const int reset_mode = ReadIniInt(ini_path, L"nuked-sc55", L"system_reset", 1);
    if (reset_mode == 1) {
        cfg.system_reset = EMU_SystemReset::GS_RESET;
    } else if (reset_mode == 2) {
        cfg.system_reset = EMU_SystemReset::GM_RESET;
    }

    const int buffer_frames = ReadIniInt(ini_path, L"nuked-sc55", L"audio_buffer_frames", 1024);
    cfg.audio_buffer_frames = static_cast<uint32_t>(std::clamp(buffer_frames, 256, 4096));

    return cfg;
}

bool CreateEmulator(HINSTANCE module) {
    const WinmmConfig cfg = LoadConfig(module);

    g_emulator = std::make_unique<Emulator>();

    EMU_Options options{};
    options.nvram_filename = cfg.nvram_file;
    if (!g_emulator->Init(options)) {
        std::fprintf(stderr, "[Nuked-SC55] Emulator::Init failed\n");
        g_emulator.reset();
        return false;
    }

    AllRomsetInfo all_info{};

    if (cfg.detect_by_filename && !DetectRomsetsByFilename(cfg.rom_dir, all_info)) {
        std::fprintf(stderr, "[Nuked-SC55] DetectRomsetsByFilename failed: %s\n", cfg.rom_dir.generic_string().c_str());
        g_emulator.reset();
        return false;
    }

    if (cfg.detect_by_hash && !DetectRomsetsByHash(cfg.rom_dir, all_info)) {
        std::fprintf(stderr, "[Nuked-SC55] DetectRomsetsByHash failed: %s\n", cfg.rom_dir.generic_string().c_str());
        g_emulator.reset();
        return false;
    }

    Romset selected = cfg.romset;
    if (!IsCompleteRomset(all_info, selected)) {
        Romset picked{};
        if (!PickCompleteRomset(all_info, picked)) {
            std::fprintf(stderr, "[Nuked-SC55] No complete romset found in: %s\n", cfg.rom_dir.generic_string().c_str());
            g_emulator.reset();
            return false;
        }
        selected = picked;
    }

    if (!LoadRomset(selected, all_info)) {
        std::fprintf(stderr, "[Nuked-SC55] LoadRomset failed for: %s\n", RomsetName(selected));
        g_emulator.reset();
        return false;
    }

    if (!g_emulator->LoadRoms(selected, all_info)) {
        std::fprintf(stderr, "[Nuked-SC55] Emulator::LoadRoms failed for: %s\n", RomsetName(selected));
        g_emulator.reset();
        return false;
    }

    g_emulator->Reset();
    g_emulator->PostSystemReset(cfg.system_reset);

    const uint32_t sample_rate = PCM_GetOutputFrequency(g_emulator->GetPCM());
    if (!g_audio.Init(sample_rate, cfg.audio_buffer_frames)) {
        std::fprintf(stderr, "[Nuked-SC55] WinMM audio init failed\n");
        g_emulator.reset();
        return false;
    }

    g_emulator->SetSampleCallback(WinmmSampleCallback, &g_audio);

    g_emulator_ready = true;
    std::fprintf(stderr,
                 "[Nuked-SC55] Emulator ready, romset=%s, sample_rate=%u, buffer_frames=%u\n",
                 RomsetName(selected),
                 sample_rate,
                 cfg.audio_buffer_frames);
    return true;
}

void WorkerLoop() {
        while (g_worker_running.load(std::memory_order_acquire)) {
        if (!g_emulator_ready || !g_emulator) {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
            continue;
        }

        g_emulator->Step();
    }
    std::fprintf(stderr, "[Nuked-SC55] Emulator stopped\n");
}

void StopRuntimeLocked() {
    g_worker_running.store(false, std::memory_order_release);
    if (g_worker_thread.joinable()) {
        g_worker_thread.join();
    }

    g_emulator_ready = false;
    g_audio.Shutdown();
    g_emulator.reset();
}

} // namespace

BOOL WINAPI NukedSC55Start() {
    std::lock_guard<std::mutex> lock(g_runtime_mutex);

    if (g_worker_running.load(std::memory_order_acquire)) {
        return TRUE;
    }

    if (!CreateEmulator(g_module)) {
        return FALSE;
    }

    g_worker_running.store(true, std::memory_order_release);
    try {
        g_worker_thread = std::thread(WorkerLoop);
    } catch (...) {
        g_worker_running.store(false, std::memory_order_release);
        g_emulator_ready = false;
        g_emulator.reset();
        std::fprintf(stderr, "[Nuked-SC55] std::thread start failed\n");
        return FALSE;
    }

    return TRUE;
}

void WINAPI NukedSC55Stop() {
    std::lock_guard<std::mutex> lock(g_runtime_mutex);
    StopRuntimeLocked();
}

Emulator emu;

UINT uMSGSDevID = -1;
HMIDIOUT hMSGSDev = nullptr;

extern "C" {
    BOOL WINAPI DllMain    (HINSTANCE, DWORD, LPVOID);
    MMRESULT midiOutOpen   (LPHMIDIOUT phmo, UINT uDeviceID, DWORD_PTR dwCallback, DWORD_PTR dwInstance, DWORD fdwOpen);
    MMRESULT midiOutClose  (HMIDIOUT hmo);
    MMRESULT midiOutLongMsg(HMIDIOUT hmo, LPMIDIHDR pmh, UINT cbmh);
    MMRESULT midiOutMessage(HMIDIOUT hmo, UINT uMsg, DWORD_PTR dw1, DWORD_PTR dw2);
}

MMRESULT midiOutOpen(LPHMIDIOUT phmo, UINT uDeviceID, DWORD_PTR dwCallback, DWORD_PTR dwInstance, DWORD fdwOpen) {
    MMRESULT result = pmidiOutOpen(phmo, uDeviceID, dwCallback, dwInstance, fdwOpen);
    std::fprintf(stderr, "[WinMM] midiOutOpen(uDeviceID = %d)\n", uDeviceID);
    if (result == MMSYSERR_NOERROR && (uDeviceID == MIDI_MAPPER || uDeviceID == uMSGSDevID)) {
        if (hMSGSDev == nullptr) {
            if (!NukedSC55Start()) {
                pmidiOutClose(*phmo);
                *phmo = nullptr;
                result = MMSYSERR_INVALPARAM;
            }
        }
        hMSGSDev = *phmo;
    }
    return result;
}

MMRESULT midiOutClose(HMIDIOUT hmo) {
    MMRESULT result = pmidiOutClose(hmo);
    if (hMSGSDev != nullptr && hMSGSDev == hmo) {
        hMSGSDev  = nullptr;
        NukedSC55Stop();
    }
    return result;
}

MMRESULT midiOutShortMsg(HMIDIOUT hmo, DWORD dwMsg) {
    MMRESULT result;
    if (hMSGSDev != nullptr && hmo == hMSGSDev) {
        uint8_t b1 = dwMsg & 0xff;
        switch (b1 & 0xf0)
        {
            case 0x80:
            case 0x90:
            case 0xa0:
            case 0xb0:
            case 0xe0:
                {
                    uint8_t buf[3] = {
                        (uint8_t)b1,
                        (uint8_t)((dwMsg >> 8) & 0xff),
                        (uint8_t)((dwMsg >> 16) & 0xff),
                    };
                    g_emulator->PostMIDI(buf);
                }
                break;
            case 0xc0:
            case 0xd0:
                {
                    uint8_t buf[2] = {
                        (uint8_t)b1,
                        (uint8_t)((dwMsg >> 8) & 0xff),
                    };
                    g_emulator->PostMIDI(buf);
                }
                break;
        }
        result = MMSYSERR_NOERROR;
    } else {
        result = pmidiOutShortMsg(hmo, dwMsg);
    }
    return result;
}

MMRESULT midiOutLongMsg(HMIDIOUT hmo, LPMIDIHDR pmh, UINT cbmh) {
    MMRESULT result;
    if (hMSGSDev != nullptr && hmo == hMSGSDev) {
        g_emulator->PostMIDI(std::span(reinterpret_cast<const uint8_t*>(pmh->lpData), pmh->dwBytesRecorded));
        result = MMSYSERR_NOERROR;
    } else {
        result = pmidiOutLongMsg(hmo, pmh, cbmh);
    }
    return result;
}

MMRESULT midiOutMessage(HMIDIOUT hmo, UINT uMsg, DWORD_PTR dw1, DWORD_PTR dw2) {
    MMRESULT result;
    if (hMSGSDev != nullptr && hmo == hMSGSDev) {
        result = MMSYSERR_NOERROR;
    } else {
        result = pmidiOutMessage(hmo, uMsg, dw1, dw2);
    }
    return result;
}

MMRESULT midiOutReset(HMIDIOUT hmo) {
    MMRESULT result;
    if (hMSGSDev != nullptr && hmo == hMSGSDev) {
        for (uint8_t i = 0; i < 16; i++) {
            g_emulator->PostMIDI(0xB0 | i);
            g_emulator->PostMIDI(0x78);
            g_emulator->PostMIDI(0x00);
            g_emulator->PostMIDI(0xB0 | i);
            g_emulator->PostMIDI(0x7B);
            g_emulator->PostMIDI(0x00);
        }
        result = MMSYSERR_NOERROR;
    } else {
        result = pmidiOutReset(hmo);
    }
    return result;
}

static inline uint16_t toLogarithm(int32_t volume) {
    return (uint16_t)((volume * volume) >> 16);
}

MMRESULT midiOutSetVolume(HMIDIOUT hmo, DWORD dwVolume) {
    MMRESULT result;
    if (hMSGSDev != nullptr && hmo == hMSGSDev) {
        g_audio.lvolume = toLogarithm((dwVolume >>  0) & 0xFFFF);
        g_audio.rvolume = toLogarithm((dwVolume >> 16) & 0xFFFF);
        result = MMSYSERR_NOERROR;
    } else {
        result = pmidiOutSetVolume(hmo, dwVolume);
    }
    return result;
}

BOOL WINAPI DllMain(HINSTANCE handle, DWORD dword, LPVOID lpvoid) {
    (void) lpvoid;
    switch(dword) {
        case DLL_PROCESS_ATTACH:
            g_module = handle;
            SetupLogOutputFromConfig(handle);
            MIDIOUTCAPSW moc;
            UINT numMidiDevs;
            char WinMMPath[MAX_PATH];
            GetSystemDirectoryA(WinMMPath, MAX_PATH);
            strcat_s(WinMMPath, MAX_PATH, "\\WinMM.dll");
            hWinMM = LoadLibraryA(WinMMPath);
            int err = GetLastError();
            if(err != NO_ERROR || hWinMM == nullptr) {
                goto fail;
            }
#define GET_PROC(name, type) \
        p##name = (type) (void *)GetProcAddress(hWinMM, #name); \
        if (!p##name) { goto fail; }
            GET_PROC(midiOutGetNumDevs,      midiOutGetNumDevsProc);
            GET_PROC(midiOutGetDevCapsW,     midiOutGetDevCapsWProc);
            GET_PROC(midiOutOpen,            midiOutOpenProc);
            GET_PROC(midiOutClose,           midiOutCloseProc);
            GET_PROC(midiOutShortMsg,        midiOutShortMsgProc);
            GET_PROC(midiOutLongMsg,         midiOutLongMsgProc);
            GET_PROC(midiOutMessage,         midiOutMessageProc);
            GET_PROC(midiOutReset,           midiOutResetProc);
            GET_PROC(midiOutSetVolume,       midiOutSetVolumeProc);
            GET_PROC(waveOutOpen,            waveOutOpenProc);
            GET_PROC(waveOutClose,           waveOutCloseProc);
            GET_PROC(waveOutPrepareHeader,   waveOutPrepareHeaderProc);
            GET_PROC(waveOutUnprepareHeader, waveOutUnprepareHeaderProc);
            GET_PROC(waveOutWrite,           waveOutWriteProc);
            GET_PROC(waveOutReset,           waveOutResetProc);
            GET_PROC(waveOutGetErrorTextA,   waveOutGetErrorTextAProc);
#undef GET_PROC
            numMidiDevs = pmidiOutGetNumDevs();
            for (UINT i = 0; i < numMidiDevs; i++) {
                pmidiOutGetDevCapsW(i, &moc, sizeof(MIDIOUTCAPSW));
                if (wcsstr(moc.szPname, L"Microsoft GS Wavetable Synth")) {
                    uMSGSDevID = i;
                    std::fprintf(stderr, "[WinMM] uMSGSDevID = %d\n", i);
                }
            }
            // MessageBoxA(NULL, "WinMM loaded.", "WinMM-Nuked-SC55", MB_ICONASTERISK);
            return TRUE;
            fail:
            if (hWinMM != nullptr)
                FreeLibrary(hWinMM);
            hWinMM = nullptr;
            MessageBoxA(NULL, "WinMM load failed.", "WinMM-Nuked-SC55", MB_ICONHAND);
            return FALSE;
    }
    if (dword == DLL_PROCESS_DETACH) {
        // FIXME: Stuck at dll unload then crash
        std::lock_guard<std::mutex> lock(g_runtime_mutex);
        StopRuntimeLocked();
        if (hWinMM != nullptr) {
            FreeLibrary(hWinMM);
            hWinMM = nullptr;
        }
    }
    return TRUE;
}
