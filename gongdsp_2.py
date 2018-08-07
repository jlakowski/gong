import numpy as np
import scipy.io.wavfile
import matplotlib.pyplot as plt
import scipy.signal as signal
import os
import glob
import scipy.fftpack as fftpack

import xlsxwriter

#wd = r'~/code/gong/'
#os.chdir(wd)

fnames = glob.glob('*.wav')
workbook = xlsxwriter.Workbook('gong_peaks.xlsx')

for i in range(len(fnames)):
    fs, sig = scipy.io.wavfile.read(fnames[i])
    
    f, t, Sxx = signal.spectrogram(sig[:,0], fs, window='hamming',nperseg=int(fs/8),noverlap=int(fs/16),nfft=int(fs/4))

    f2, Pxx_spec = signal.welch(sig[:,0], fs, 'hamming', nperseg=fs/8, noverlap=fs/16, scaling='spectrum')
    #Sxxdb = 20*np.log10(np.absolute(Sxx)/pref)
    
    #plot the spectrogram
    plt.figure()
    fig = plt.gcf() #oh get current fig!
    fig.set_size_inches(18.5,10)
    plt.pcolormesh(t, f,np.log(Sxx+1))
    plt.ylabel('Frequency [Hz]')
    plt.xlabel('Time [sec]')
    plt.title(fnames[i])
    plt.ylim((0, 7000))
    plt.savefig(fnames[i] + '_spectrogram.png', dpi=100, bbox_inches='tight')
    
    plt.close()
    
    #find the  amplitude / frequency of peaks 
    #then write to a spreadsheet
    worksheet = workbook.add_worksheet(name=fnames[i])
    row = 0
    col = 0
    
 
    worksheet.write(row,0,'Amplitude')
    worksheet.write(row,1,'Frequency')
    peaks, _ = signal.find_peaks(Pxx_spec)
    peak_coords = np.zeros((len(peaks),2))
    num_peaks = len(peaks)
    if len(peaks) > 100:
        num_peaks = 100 #only publish the first 100
    for j in range(num_peaks):
        row = j+1
        peak_freq = f2[peaks[j]]
        peak_amp = Pxx_spec[peaks[j]]
        peak_coords[j,:] = np.array([peak_freq,peak_amp])
        worksheet.write(row,0,peak_freq)
        worksheet.write(row,1,peak_amp)
        
    
    
    
    plt.figure()
    fig = plt.gcf() #oh get current fig!
    fig.set_size_inches(18.5,10)
    Pxx_db = 20*np.log10(Pxx_spec)
    plt.plot(f2, Pxx_db)
    plt.xlabel('frequency [Hz]')
    plt.ylabel('Spectrum [dB RMS]')
    plt.xlim((0,7000))
    plt.title(fnames[i])
    #plt.show()
    plt.savefig(fnames[i] + '_db_psd.png', dpi=100,bbox_inches='tight')

    plt.close()



workbook.close()
